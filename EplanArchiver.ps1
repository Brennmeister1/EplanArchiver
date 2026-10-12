# ============================================================================
# EPLAN Projekt-/Benutzerliste
# ============================================================================
# Durchsucht ein Verzeichnis nach EPLAN-Projekten (.elk/.edb und .zw1),
# liest die ProjectInfo.xml und sammelt Projektname + Benutzer-E-Mail-Adressen.
# Die Properties werden über ihre ID gelesen -> funktioniert für deutsche und
# englische Projekte gleichermaßen.
# Ergebnis: geordnete Listen als CSV (Excel-kompatibel, Trennzeichen ';').
#
# Es werden KEINE Projekte verschoben, kopiert oder verändert (nur Lesezugriff).
#
# Version: 2.0
# ============================================================================

#Requires -Version 5.1

[CmdletBinding()]
param(
    # Quellverzeichnis für die Suche
    [string]$SourcePath = "C:\EPLAN_Projekte",

    # Ausgabeverzeichnis für die Listen
    [string]$OutputPath = "C:\Temp",

    # Projekte ohne gefundene E-Mail-Adresse trotzdem aufführen
    [bool]$IncludeProjectsWithoutEmail = $true
)

# ============================================================================
# KONFIGURATION
# ============================================================================

$Config = @{
    SourcePath                  = $SourcePath
    IncludeProjectsWithoutEmail = $IncludeProjectsWithoutEmail

    # Liste 1: eine Zeile pro Projekt + E-Mail, sortiert nach Projekt, dann E-Mail
    CsvByProjectPath = Join-Path $OutputPath "EPLAN_Projekte_Benutzer.csv"

    # Liste 2: eine Zeile pro E-Mail mit allen zugehörigen Projekten
    CsvByUserPath    = Join-Path $OutputPath "EPLAN_Benutzer_Projekte.csv"

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

    if (-not (Test-Path $Config.SourcePath)) {
        Write-LogMessage "FEHLER: Quellverzeichnis '$($Config.SourcePath)' existiert nicht!" "ERROR"
        return $false
    }
    return $true
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

function Find-EplanProjects {
    param([string]$SearchPath)

    Write-LogMessage "Durchsuche Verzeichnis: $SearchPath"
    $Results = New-Object System.Collections.Generic.List[object]

    # --- Normale Projekte (.elk + .edb\ProjektInfo.xml) ---
    $ElkFiles = Get-ChildItem -Path $SearchPath -Filter "*.elk" -Recurse -File -ErrorAction SilentlyContinue
    foreach ($ElkFile in $ElkFiles) {
        $ProjectName = [System.IO.Path]::GetFileNameWithoutExtension($ElkFile.Name)
        $EdbDir      = Join-Path $ElkFile.Directory.FullName "$ProjectName.edb"
        $XmlPath     = $InfoFileNames |
            ForEach-Object { Join-Path $EdbDir $_ } |
            Where-Object { Test-Path -LiteralPath $_ } |
            Select-Object -First 1

        if (-not $XmlPath) {
            Write-LogMessage "ProjectInfo.xml nicht gefunden für '$ProjectName' in $EdbDir" "WARN"
            $Results.Add(@{ ProjectName = $ProjectName; Type = "Projekt"; Path = $ElkFile.FullName; Info = $null; Error = "ProjectInfo.xml fehlt" })
            continue
        }

        try {
            $Info = Get-ProjectInfoFromText (Read-InfoFile $XmlPath)
            Write-LogMessage "Projekt '$ProjectName': $($Info.EditorName) <$($Info.EditorEmail)>"
            $Results.Add(@{ ProjectName = $ProjectName; Type = "Projekt"; Path = $ElkFile.FullName; Info = $Info; Error = "" })
        }
        catch {
            Write-LogMessage "Fehler beim Lesen von '$XmlPath': $_" "ERROR"
            $Results.Add(@{ ProjectName = $ProjectName; Type = "Projekt"; Path = $ElkFile.FullName; Info = $null; Error = $_.Exception.Message })
        }
    }

    # --- Projektarchive (.zw1) ---
    $ZwFiles = Get-ChildItem -Path $SearchPath -Filter "*.zw1" -Recurse -File -ErrorAction SilentlyContinue
    foreach ($ZwFile in $ZwFiles) {
        $ProjectName = [System.IO.Path]::GetFileNameWithoutExtension($ZwFile.Name)
        try {
            $Text = Read-InfoFromZip $ZwFile.FullName
            if (-not $Text) {
                Write-LogMessage "ProjectInfo.xml nicht im Archiv '$ProjectName' gefunden" "WARN"
                $Results.Add(@{ ProjectName = $ProjectName; Type = "Archiv"; Path = $ZwFile.FullName; Info = $null; Error = "ProjectInfo.xml fehlt im Archiv" })
                continue
            }
            $Info = Get-ProjectInfoFromText $Text
            Write-LogMessage "Archiv '$ProjectName': $($Info.EditorName) <$($Info.EditorEmail)>"
            $Results.Add(@{ ProjectName = $ProjectName; Type = "Archiv"; Path = $ZwFile.FullName; Info = $Info; Error = "" })
        }
        catch {
            Write-LogMessage "Archiv '$($ZwFile.FullName)' nicht lesbar: $_" "WARN"
            $Results.Add(@{ ProjectName = $ProjectName; Type = "Archiv"; Path = $ZwFile.FullName; Info = $null; Error = $_.Exception.Message })
        }
    }

    Write-LogMessage "Gefunden: $($Results.Count) Projekte/Archive"
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
            'Projektname'       = $P.ProjectName
            'Letzter Bearbeiter'= if ($I) { $I.EditorName }  else { "" }
            'E-Mail'            = if ($I) { $I.EditorEmail } else { "" }
            'Anmeldename'       = if ($I) { $I.EditorLogin } else { "" }
            'Ersteller'         = if ($I) { $I.Creator }     else { "" }
            'Erstellt'          = if ($I) { Format-Date $I.CreationDate }     else { "" }
            'Letzte Änderung'   = if ($I) { Format-Date $I.ModificationDate } else { "" }
            'Typ'               = $P.Type
            'Pfad'              = $P.Path
            'Hinweis'           = $Hint
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
    if (-not $Rows) {
        Write-LogMessage "Keine Daten für '$Path'." "WARN"
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
    Write-LogMessage "  Quellpfad: $($Config.SourcePath)"

    if (-not (Initialize-Directories)) {
        Write-LogMessage "Initialisierung fehlgeschlagen. Abbruch." "ERROR"
        return
    }

    $Projects = Find-EplanProjects -SearchPath $Config.SourcePath
    if ($Projects.Count -eq 0) {
        Write-LogMessage "Keine EPLAN-Projekte gefunden."
        return
    }

    $ByProject = New-ProjectUserList -Projects $Projects
    $ByUser    = New-UserProjectList -ProjectRows $ByProject

    Export-List -Rows $ByProject -Path $Config.CsvByProjectPath
    Export-List -Rows $ByUser    -Path $Config.CsvByUserPath

    $WithMail = @($Projects | Where-Object { $_.Info -and $_.Info.EditorEmail }).Count
    Write-LogMessage "ZUSAMMENFASSUNG:"
    Write-LogMessage "  Projekte gesamt:        $($Projects.Count)"
    Write-LogMessage "  davon mit E-Mail:       $WithMail"
    Write-LogMessage "  Eindeutige E-Mails:     $(@($ByUser).Count)"
    Write-LogMessage "EPLAN Projekt-/Benutzerliste beendet"
}

Start-EplanUserCollector

# ============================================================================
# VERWENDUNG
# ============================================================================
#   .\EplanArchiver.ps1
#   .\EplanArchiver.ps1 -SourcePath "\\server\EPLAN\Projekte" -OutputPath "C:\Temp"
#
# Ausgabe:
#   EPLAN_Projekte_Benutzer.csv  - Projekt | Letzter Bearbeiter | E-Mail | ... (sortiert nach Projekt)
#   EPLAN_Benutzer_Projekte.csv  - E-Mail | Name | Anzahl | Projekte (sortiert nach E-Mail)
#
# Die Werte stammen aus der ProjectInfo.xml (Property-IDs siehe $PropId oben).
# Die E-Mail ist die des letzten Bearbeiters (ID 3013); für den Ersteller
# (ID 10020) speichert EPLAN nur die Kennung, keine E-Mail.
#
# Nur Lesezugriff, keine externen Abhängigkeiten (PowerShell 5.1+).
# ============================================================================
