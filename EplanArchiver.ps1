# ============================================================================
# EPLAN Projekt-/Benutzerliste
# ============================================================================
# Durchsucht ein oder mehrere Verzeichnisse nach EPLAN-Projekten (.elk/.edb
# und .zw1), liest die ProjectInfo.xml und sammelt Projektname + Benutzer-
# E-Mail-Adressen. Die Properties werden über ihre ID gelesen -> funktioniert
# für deutsche und englische Projekte gleichermaßen.
#
# Ergebnis: geordnete Listen als CSV (Excel-kompatibel, Trennzeichen ';').
# Ausnahmen (Pfade und Projektnummern) werden in einer Textdatei gepflegt.
#
# Es werden KEINE Projekte verschoben, kopiert oder verändert (nur Lesezugriff).
#
# Version: 2.1
# ============================================================================

#Requires -Version 5.1

[CmdletBinding()]
param(
    # Verzeichnis(se), in denen gesucht wird. Mehrere mit Komma oder ';' trennen.
    [Alias('Pfad', 'SourcePath')]
    [string[]]$SearchPath = @("C:\EPLAN_Projekte"),

    # Ausgabeverzeichnis für Listen und Log
    [string]$OutputPath = "C:\Temp",

    # Ausnahmeliste (Pfade und Projektnummern, eine Zeile pro Eintrag)
    [string]$ExclusionFile = (Join-Path $PSScriptRoot "EPLAN_Ausnahmen.txt"),

    # Dry-Run: nur anzeigen, keine CSV-Dateien schreiben
    [switch]$DryRun,

    # Projekte ohne hinterlegte E-Mail-Adresse trotzdem aufführen
    [bool]$IncludeProjectsWithoutEmail = $true
)

# ============================================================================
# KONFIGURATION
# ============================================================================

$Config = @{
    # Mehrere Pfade per Komma (in der Konsole) oder ';' (z. B. Aufgabenplanung)
    SearchPath                  = @($SearchPath | ForEach-Object { $_ -split ';' } |
                                    ForEach-Object { $_.Trim() } | Where-Object { $_ })
    ExclusionFile               = $ExclusionFile
    DryRun                      = [bool]$DryRun
    IncludeProjectsWithoutEmail = $IncludeProjectsWithoutEmail

    # Liste 1: eine Zeile pro Projekt, sortiert nach Projektname
    CsvByProjectPath = Join-Path $OutputPath "EPLAN_Projekte_Benutzer.csv"

    # Liste 2: eine Zeile pro E-Mail mit allen zugehörigen Projekten
    CsvByUserPath    = Join-Path $OutputPath "EPLAN_Benutzer_Projekte.csv"

    # Liste 3: alles, was über die Ausnahmeliste übersprungen wurde
    CsvExcludedPath  = Join-Path $OutputPath "EPLAN_Ausgenommen.csv"

    LogPath          = Join-Path $OutputPath "EPLAN_Archiver.log"
}

# Dateinamen der Projektinfo (beide Schreibweisen werden akzeptiert)
$InfoFileNames = @('ProjectInfo.xml', 'ProjektInfo.xml')

# EPLAN Property-IDs (sprachunabhängig, gleiche IDs in DE und EN)
$PropId = @{
    EditorName   = '3011'   # Letzter Bearbeiter: Name        / Last editor: Name
    EditorEmail  = '3013'   # Letzter Bearbeiter: E-Mail      / Last editor: E-mail
    ProjectPath  = '10009'  # Projektname (vollständig)       / Project name (full)
    Creator      = '10020'  # Ersteller                       / Creator
    CreationDate = '10021'  # Erstellungsdatum                / Creation date
    EditorLogin  = '10022'  # Letzter Bearbeiter: Anmeldename / Last editor: Sign-in name
    ModDate      = '10023'  # Änderungsdatum                  / Modification date
}

# Vorlage für eine neue Ausnahmeliste
$ExclusionTemplate = @'
# ============================================================================
# EPLAN Ausnahmeliste
# ============================================================================
# Eine Ausnahme pro Zeile. Leere Zeilen und Zeilen mit # werden ignoriert.
#
# PFADE (alles, was einen \ oder / enthält):
#   Der Ordner und alle Unterordner werden komplett übersprungen.
#   Platzhalter * und ? sind erlaubt.
#     J:\EPLAN\Projects\Prod\TestProjects
#     J:\EPLAN\Projects\*\Archiv
#     \\server\freigabe\EPLAN\Alt
#
# PROJEKTNUMMERN (alles ohne \ oder /):
#   Wird mit dem Projektnamen (Dateiname ohne .elk/.zw1) verglichen,
#   Groß-/Kleinschreibung egal. Platzhalter * und ? sind erlaubt.
#     K998Y72
#     K998*
#     Neues Projekt*
# ============================================================================

'@

# ============================================================================
# HILFSFUNKTIONEN
# ============================================================================

function Write-LogMessage {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )

    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry  = "[$TimeStamp] [$Level] $Message"

    switch ($Level) {
        "ERROR" { Write-Host $LogEntry -ForegroundColor Red }
        "WARN"  { Write-Host $LogEntry -ForegroundColor Yellow }
        "SKIP"  { Write-Host $LogEntry -ForegroundColor DarkGray }
        default { Write-Host $LogEntry -ForegroundColor Green }
    }

    try {
        Add-Content -Path $Config.LogPath -Value $LogEntry -Encoding UTF8
    }
    catch {
        Write-Host "Fehler beim Schreiben in Log-Datei: $_" -ForegroundColor Red
    }
}

function Initialize-Directories {
    $OutDir = Split-Path $Config.CsvByProjectPath
    if (-not (Test-Path $OutDir)) {
        New-Item -Path $OutDir -ItemType Directory -Force | Out-Null
    }

    $Ok = $true
    foreach ($P in $Config.SearchPath) {
        if (-not (Test-Path -LiteralPath $P -PathType Container)) {
            Write-LogMessage "FEHLER: Suchpfad '$P' existiert nicht!" "ERROR"
            $Ok = $false
        }
    }
    return $Ok
}

# Normalisiert einen Pfad für Vergleiche (Trennzeichen, kein abschließendes \)
function Format-PathKey {
    param([string]$Path)
    return ($Path -replace '[\\/]+', '\').TrimEnd('\')
}

# Liest die Ausnahmeliste; legt eine Vorlage an, falls sie nicht existiert
function Read-Exclusions {
    $Result = @{ Paths = @(); Projects = @() }

    if (-not $Config.ExclusionFile) { return $Result }

    if (-not (Test-Path -LiteralPath $Config.ExclusionFile)) {
        try {
            Set-Content -LiteralPath $Config.ExclusionFile -Value $ExclusionTemplate -Encoding UTF8
            Write-LogMessage "Ausnahmeliste nicht gefunden - Vorlage angelegt: $($Config.ExclusionFile)" "WARN"
        }
        catch {
            Write-LogMessage "Ausnahmeliste '$($Config.ExclusionFile)' fehlt und konnte nicht angelegt werden: $_" "WARN"
        }
        return $Result
    }

    foreach ($Line in Get-Content -LiteralPath $Config.ExclusionFile -Encoding UTF8) {
        $Entry = $Line.Trim()
        if (-not $Entry -or $Entry.StartsWith('#')) { continue }

        if ($Entry -match '[\\/]') { $Result.Paths    += (Format-PathKey $Entry) }
        else                       { $Result.Projects += $Entry }
    }

    Write-LogMessage "Ausnahmeliste: $($Result.Paths.Count) Pfad(e), $($Result.Projects.Count) Projektnummer(n) - $($Config.ExclusionFile)"
    foreach ($P in $Result.Paths)    { Write-LogMessage "  Pfad ausgenommen:    $P" }
    foreach ($P in $Result.Projects) { Write-LogMessage "  Projekt ausgenommen: $P" }
    return $Result
}

# Prüft, ob ein Pfad (Ordner oder Datei) unter eine Pfad-Ausnahme fällt.
# Liefert die greifende Regel oder $null.
function Test-PathExcluded {
    param([string]$Path, [string[]]$Rules)

    $Key = Format-PathKey $Path
    foreach ($Rule in $Rules) {
        if ($Rule -match '[\*\?]') {
            # Platzhalter: der Pfad selbst oder ein übergeordneter Ordner passt
            if ($Key -like $Rule -or $Key -like "$Rule\*") { return $Rule }
        }
        elseif ($Key -eq $Rule -or $Key.StartsWith("$Rule\", [System.StringComparison]::OrdinalIgnoreCase)) {
            return $Rule
        }
    }
    return $null
}

# Prüft, ob ein Projektname unter eine Projektnummer-Ausnahme fällt
function Test-ProjectExcluded {
    param([string]$ProjectName, [string[]]$Rules)

    foreach ($Rule in $Rules) {
        if ($ProjectName -like $Rule) { return $Rule }
    }
    return $null
}

function ConvertFrom-UnixTimestamp {
    param([string]$UnixTimestamp)

    try {
        # EPLAN speichert UTC -> in lokale Zeit umrechnen
        return [DateTimeOffset]::FromUnixTimeSeconds([long]$UnixTimestamp).LocalDateTime
    }
    catch {
        return $null
    }
}

# Liest den Wert einer Property über ihre ID (index="0").
# Bewusst per Regex statt XML-Parser: robust gegen nicht wohlgeformte Inhalte
# (z. B. unmaskierte "<1100>"-Werte) und unabhängig von der Sprache.
function Get-PropValue {
    param([string]$Text, [string]$Id)

    $Pattern = '<Property\b(?=[^>]*\bid="' + $Id + '")(?=[^>]*\bindex="0")[^>]*?(?:/>|>(?<v>.*?)</Property>)'
    $M = [regex]::Match($Text, $Pattern, 'Singleline')
    if ($M.Success -and $M.Groups['v'].Success) {
        return [System.Net.WebUtility]::HtmlDecode($M.Groups['v'].Value).Trim()
    }
    return ""
}

# Wertet den Inhalt einer ProjectInfo.xml aus
function Get-ProjectInfoFromText {
    param([string]$Text)

    if ($Text -notmatch '<Property\b') { throw "Keine EPLAN-Properties in der Datei gefunden" }

    $Created  = Get-PropValue $Text $PropId.CreationDate
    $Modified = Get-PropValue $Text $PropId.ModDate

    return @{
        EditorName       = Get-PropValue $Text $PropId.EditorName
        EditorEmail      = (Get-PropValue $Text $PropId.EditorEmail).ToLowerInvariant()
        EditorLogin      = Get-PropValue $Text $PropId.EditorLogin
        Creator          = Get-PropValue $Text $PropId.Creator
        ProjectPath      = Get-PropValue $Text $PropId.ProjectPath
        CreationDate     = if ($Created)  { ConvertFrom-UnixTimestamp $Created }  else { $null }
        ModificationDate = if ($Modified) { ConvertFrom-UnixTimestamp $Modified } else { $null }
    }
}

# Liest eine Datei als Text (Encoding per BOM erkannt, sonst UTF-8)
function Read-InfoFile {
    param([string]$Path)
    return [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8)
}

# Liest die ProjectInfo.xml aus einem .zw1-Archiv (ZIP) als Text
function Read-InfoFromZip {
    param([string]$ArchivePath)

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $Zip = [System.IO.Compression.ZipFile]::OpenRead($ArchivePath)
    try {
        $Entry = $Zip.Entries | Where-Object { $InfoFileNames -contains $_.Name } | Select-Object -First 1
        if (-not $Entry) { return $null }

        $Reader = New-Object System.IO.StreamReader($Entry.Open(), [System.Text.Encoding]::UTF8, $true)
        try { return $Reader.ReadToEnd() }
        finally { $Reader.Dispose() }
    }
    finally { $Zip.Dispose() }
}

# ============================================================================
# HAUPTFUNKTIONEN
# ============================================================================

# Durchläuft die Verzeichnisse selbst, damit ausgenommene Ordner gar nicht
# erst betreten werden (wichtig bei großen Netzlaufwerken). In .edb-Ordner
# wird ebenfalls nicht hineingesucht.
function Get-EplanFiles {
    param([string]$Root, $Exclusions, $Excluded)

    $Files = New-Object System.Collections.Generic.List[System.IO.FileInfo]
    $Stack = New-Object System.Collections.Generic.Stack[System.IO.DirectoryInfo]
    $Stack.Push((Get-Item -LiteralPath $Root))

    while ($Stack.Count -gt 0) {
        $Dir = $Stack.Pop()

        $Rule = Test-PathExcluded $Dir.FullName $Exclusions.Paths
        if ($Rule) {
            Write-LogMessage "Ordner übersprungen (Ausnahme '$Rule'): $($Dir.FullName)" "SKIP"
            $Excluded.Add([PSCustomObject]@{ 'Projektname' = ""; 'Pfad' = $Dir.FullName; 'Art' = "Ordner"; 'Regel' = $Rule })
            continue
        }

        try {
            foreach ($F in $Dir.EnumerateFiles()) {
                if ($F.Extension -eq '.elk' -or $F.Extension -eq '.zw1') { $Files.Add($F) }
            }
            foreach ($Sub in $Dir.EnumerateDirectories()) {
                if ($Sub.Name -notlike '*.edb') { $Stack.Push($Sub) }
            }
        }
        catch {
            Write-LogMessage "Kein Zugriff auf '$($Dir.FullName)': $($_.Exception.Message)" "WARN"
        }
    }
    return $Files
}

function Read-EplanProject {
    param([System.IO.FileInfo]$File)

    $ProjectName = [System.IO.Path]::GetFileNameWithoutExtension($File.Name)
    $Type        = if ($File.Extension -eq '.zw1') { "Archiv" } else { "Projekt" }
    $Result      = @{ ProjectName = $ProjectName; Type = $Type; Path = $File.FullName; Info = $null; Error = "" }

    try {
        if ($Type -eq "Projekt") {
            $EdbDir  = Join-Path $File.Directory.FullName "$ProjectName.edb"
            $XmlPath = $InfoFileNames |
                ForEach-Object { Join-Path $EdbDir $_ } |
                Where-Object { Test-Path -LiteralPath $_ } |
                Select-Object -First 1

            if (-not $XmlPath) {
                $Result.Error = "ProjectInfo.xml fehlt"
                Write-LogMessage "ProjectInfo.xml nicht gefunden für '$ProjectName' in $EdbDir" "WARN"
                return $Result
            }
            $Text = Read-InfoFile $XmlPath
        }
        else {
            $Text = Read-InfoFromZip $File.FullName
            if (-not $Text) {
                $Result.Error = "ProjectInfo.xml fehlt im Archiv"
                Write-LogMessage "ProjectInfo.xml nicht im Archiv '$ProjectName' gefunden" "WARN"
                return $Result
            }
        }

        $Result.Info = Get-ProjectInfoFromText $Text
        Write-LogMessage "$Type '$ProjectName': $($Result.Info.EditorName) <$($Result.Info.EditorEmail)>"
    }
    catch {
        $Result.Error = $_.Exception.Message
        Write-LogMessage "$Type '$($File.FullName)' nicht lesbar: $($_.Exception.Message)" "WARN"
    }
    return $Result
}

function Find-EplanProjects {
    param($Exclusions, $Excluded)

    $Results = New-Object System.Collections.Generic.List[object]
    $Seen    = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($Root in $Config.SearchPath) {
        Write-LogMessage "Durchsuche Verzeichnis: $Root"

        foreach ($File in (Get-EplanFiles -Root $Root -Exclusions $Exclusions -Excluded $Excluded)) {
            if (-not $Seen.Add($File.FullName)) { continue }   # bei überlappenden Suchpfaden

            $ProjectName = [System.IO.Path]::GetFileNameWithoutExtension($File.Name)

            # Einzelne Datei per Pfad ausgenommen?
            $Rule = Test-PathExcluded $File.FullName $Exclusions.Paths
            if (-not $Rule) { $Rule = Test-ProjectExcluded $ProjectName $Exclusions.Projects }
            if ($Rule) {
                Write-LogMessage "Projekt übersprungen (Ausnahme '$Rule'): $($File.FullName)" "SKIP"
                $Excluded.Add([PSCustomObject]@{ 'Projektname' = $ProjectName; 'Pfad' = $File.FullName; 'Art' = "Projekt"; 'Regel' = $Rule })
                continue
            }

            $Results.Add((Read-EplanProject $File))
        }
    }

    Write-LogMessage "Gefunden: $($Results.Count) Projekte/Archive, $($Excluded.Count) Ausnahme(n) übersprungen"
    return $Results
}

function Format-Date {
    param($Date)
    if ($Date) { return $Date.ToString('dd.MM.yyyy HH:mm') }
    return ""
}

# Liste 1: eine Zeile pro Projekt, sortiert nach Projektname
function New-ProjectUserList {
    param($Projects)

    $Rows = foreach ($P in $Projects) {
        $I = $P.Info
        if (-not $Config.IncludeProjectsWithoutEmail -and (-not $I -or -not $I.EditorEmail)) { continue }

        $Hint = $P.Error
        if (-not $Hint -and -not $I.EditorEmail) { $Hint = "Keine E-Mail hinterlegt" }

        [PSCustomObject]@{
            'Projektname'        = $P.ProjectName
            'Letzter Bearbeiter' = if ($I) { $I.EditorName }  else { "" }
            'E-Mail'             = if ($I) { $I.EditorEmail } else { "" }
            'Anmeldename'        = if ($I) { $I.EditorLogin } else { "" }
            'Ersteller'          = if ($I) { $I.Creator }     else { "" }
            'Erstellt'           = if ($I) { Format-Date $I.CreationDate }     else { "" }
            'Letzte Änderung'    = if ($I) { Format-Date $I.ModificationDate } else { "" }
            'Typ'                = $P.Type
            'Pfad'               = $P.Path
            'Hinweis'            = $Hint
        }
    }

    return @($Rows | Sort-Object 'Projektname', 'Pfad')
}

# Liste 2: eine Zeile pro E-Mail mit allen zugehörigen Projekten
function New-UserProjectList {
    param($ProjectRows)

    $ProjectRows |
        Where-Object { $_.'E-Mail' } |
        Group-Object 'E-Mail' |
        Sort-Object Name |
        ForEach-Object {
            $Names = @($_.Group | Select-Object -ExpandProperty 'Projektname' | Sort-Object -Unique)
            $User  = ($_.Group | Where-Object { $_.'Letzter Bearbeiter' } | Select-Object -First 1).'Letzter Bearbeiter'
            [PSCustomObject]@{
                'E-Mail'   = $_.Name
                'Name'     = $User
                'Anzahl'   = $Names.Count
                'Projekte' = ($Names -join ', ')
            }
        }
}

function Export-List {
    param($Rows, [string]$Path)

    if ($Config.DryRun) {
        Write-LogMessage "[DRY-RUN] Würde $(@($Rows).Count) Zeile(n) schreiben: $Path"
        return
    }
    if (-not $Rows) {
        # Leere Liste trotzdem schreiben, damit keine veralteten Daten liegen bleiben
        Set-Content -Path $Path -Value "" -Encoding UTF8
        Write-LogMessage "Liste leer: $Path" "WARN"
        return
    }
    $Rows | Export-Csv -Path $Path -Delimiter ';' -Encoding UTF8 -NoTypeInformation
    Write-LogMessage "Liste erstellt: $Path"
}

# ============================================================================
# HAUPTPROGRAMM
# ============================================================================

function Start-EplanUserCollector {
    Write-LogMessage "EPLAN Projekt-/Benutzerliste gestartet"
    Write-LogMessage "  Suchpfad(e):   $($Config.SearchPath -join '; ')"
    Write-LogMessage "  Ausnahmeliste: $($Config.ExclusionFile)"
    Write-LogMessage "  Dry-Run:       $($Config.DryRun)"

    if (-not (Initialize-Directories)) {
        Write-LogMessage "Initialisierung fehlgeschlagen. Abbruch." "ERROR"
        return
    }

    $Exclusions = Read-Exclusions
    $Excluded   = New-Object System.Collections.Generic.List[object]

    $Projects = Find-EplanProjects -Exclusions $Exclusions -Excluded $Excluded

    $ByProject = New-ProjectUserList -Projects $Projects
    $ByUser    = @(New-UserProjectList -ProjectRows $ByProject)
    $ExclRows  = @($Excluded | Sort-Object 'Art', 'Projektname', 'Pfad')

    Export-List -Rows $ByProject -Path $Config.CsvByProjectPath
    Export-List -Rows $ByUser    -Path $Config.CsvByUserPath
    Export-List -Rows $ExclRows  -Path $Config.CsvExcludedPath

    if ($Config.DryRun) {
        Write-Host ""
        Write-Host "=== DRY-RUN: Projekte ===" -ForegroundColor Cyan
        $ByProject | Format-Table 'Projektname', 'Letzter Bearbeiter', 'E-Mail', 'Letzte Änderung', 'Hinweis' -AutoSize | Out-String -Width 4096 | Write-Host
        Write-Host "=== DRY-RUN: Benutzer ===" -ForegroundColor Cyan
        $ByUser | Format-Table -AutoSize | Out-String -Width 4096 | Write-Host
        if ($ExclRows) {
            Write-Host "=== DRY-RUN: Ausgenommen ===" -ForegroundColor Cyan
            $ExclRows | Format-Table -AutoSize | Out-String -Width 4096 | Write-Host
        }
    }

    $WithMail = @($Projects | Where-Object { $_.Info -and $_.Info.EditorEmail }).Count
    Write-LogMessage "ZUSAMMENFASSUNG:"
    Write-LogMessage "  Projekte gesamt:        $($Projects.Count)"
    Write-LogMessage "  davon mit E-Mail:       $WithMail"
    Write-LogMessage "  Eindeutige E-Mails:     $($ByUser.Count)"
    Write-LogMessage "  Ausgenommen:            $($ExclRows.Count)"
    if ($Config.DryRun) {
        Write-LogMessage "Dies war ein DRY-RUN. Es wurden keine Listen geschrieben."
    }
    Write-LogMessage "EPLAN Projekt-/Benutzerliste beendet"
}

Start-EplanUserCollector

# ============================================================================
# VERWENDUNG
# ============================================================================
#   Testlauf (nur Bildschirmausgabe, keine CSV-Dateien):
#     .\EplanArchiver.ps1 -Pfad "J:\EPLAN\Projects" -DryRun
#
#   Echter Lauf:
#     .\EplanArchiver.ps1 -Pfad "J:\EPLAN\Projects"
#
#   Aus der Aufgabenplanung (mehrere Pfade mit ';' trennen):
#     powershell.exe -ExecutionPolicy Bypass -File C:\powershell\EplanArchiver.ps1 -Pfad "J:\EPLAN\Projects;\\server\EPLAN"
#
#   Mehrere Suchpfade, eigene Ausnahmeliste, anderes Ausgabeverzeichnis:
#     .\EplanArchiver.ps1 -Pfad "J:\EPLAN\Projects","\\server\EPLAN" `
#                         -ExclusionFile "C:\powershell\Ausnahmen.txt" `
#                         -OutputPath "C:\Temp"
#
# Ausnahmeliste:
#   Standard ist EPLAN_Ausnahmen.txt neben dem Skript. Fehlt sie, wird beim
#   ersten Lauf eine Vorlage mit Erklärung angelegt.
#   - Zeilen mit \ oder /  -> Pfad (Ordner inkl. Unterordner wird übersprungen)
#   - alle anderen Zeilen  -> Projektnummer (Vergleich mit dem Dateinamen)
#   - Platzhalter * und ? sind erlaubt, # leitet einen Kommentar ein
#
# Ausgabe:
#   EPLAN_Projekte_Benutzer.csv  - Projekt | Letzter Bearbeiter | E-Mail | ... (sortiert nach Projekt)
#   EPLAN_Benutzer_Projekte.csv  - E-Mail | Name | Anzahl | Projekte (sortiert nach E-Mail)
#   EPLAN_Ausgenommen.csv        - was durch die Ausnahmeliste übersprungen wurde
#
# Die E-Mail ist die des letzten Bearbeiters (Property-ID 3013); für den
# Ersteller (ID 10020) speichert EPLAN nur die Kennung, keine E-Mail.
#
# Nur Lesezugriff, keine externen Abhängigkeiten (PowerShell 5.1+).
# ============================================================================
