# ============================================================================
# EPLAN Projekt-/Benutzerliste
# ============================================================================
# Durchsucht ein Verzeichnis nach EPLAN-Projekten (.elk/.edb und .zw1),
# liest die ProjektInfo.xml und sammelt Projektname + Benutzer-E-Mail-Adressen.
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

# Muster für E-Mail-Adressen
$EmailRegex = [regex]'(?i)[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}'

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
        return ([datetime]'1970-01-01 00:00:00').AddSeconds([long]$UnixTimestamp)
    }
    catch {
        return $null
    }
}

# Liefert den Property-Namen eines Knotens (Attribut name/id), falls vorhanden
function Get-PropertyLabel {
    param([System.Xml.XmlNode]$Node)

    foreach ($AttrName in @('name', 'Name', 'id', 'ID')) {
        if ($Node.Attributes -and $Node.Attributes[$AttrName]) {
            return $Node.Attributes[$AttrName].Value
        }
    }
    return $Node.LocalName
}

# Wertet den XML-Inhalt einer ProjektInfo.xml aus:
# alle E-Mail-Adressen (aus Elementtexten und Attributen) + Änderungsdatum
function Get-ProjectInfoFromXmlContent {
    param(
        [xml]$XmlContent,
        [string]$ProjectName
    )

    $Emails = @{}   # Key = E-Mail (klein), Value = Liste der Fundstellen

    foreach ($Node in $XmlContent.SelectNodes("//*")) {
        # Nur direkter Text des Knotens (nicht der Kinder)
        $OwnText = ($Node.ChildNodes |
            Where-Object { $_.NodeType -eq 'Text' -or $_.NodeType -eq 'CDATA' } |
            ForEach-Object { $_.Value }) -join ' '

        $Sources = @()
        if ($OwnText) { $Sources += , @($OwnText, (Get-PropertyLabel $Node)) }
        if ($Node.Attributes) {
            foreach ($Attr in $Node.Attributes) {
                $Sources += , @($Attr.Value, "$(Get-PropertyLabel $Node)@$($Attr.Name)")
            }
        }

        foreach ($Src in $Sources) {
            foreach ($Match in $EmailRegex.Matches($Src[0])) {
                $Mail = $Match.Value.ToLowerInvariant().TrimEnd('.')
                if (-not $Emails.ContainsKey($Mail)) { $Emails[$Mail] = New-Object System.Collections.Generic.List[string] }
                if (-not $Emails[$Mail].Contains($Src[1])) { $Emails[$Mail].Add($Src[1]) }
            }
        }
    }

    $ModificationDate = $null
    $ModNode = $XmlContent.SelectSingleNode("//Property[@name='Modification date']")
    if ($ModNode) { $ModificationDate = ConvertFrom-UnixTimestamp $ModNode.InnerText }

    return @{
        Emails           = $Emails
        ModificationDate = $ModificationDate
    }
}

function Read-XmlFile {
    param([string]$Path)
    # Encoding aus der XML-Deklaration respektieren
    $Doc = New-Object System.Xml.XmlDocument
    $Doc.Load($Path)
    return $Doc
}

function Read-XmlFromZip {
    param([string]$ArchivePath)

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $Zip = [System.IO.Compression.ZipFile]::OpenRead($ArchivePath)
    try {
        $Entry = $Zip.Entries | Where-Object { $_.Name -eq "ProjektInfo.xml" } | Select-Object -First 1
        if (-not $Entry) { return $null }

        $Stream = $Entry.Open()
        try {
            $Doc = New-Object System.Xml.XmlDocument
            $Doc.Load($Stream)
            return $Doc
        }
        finally { $Stream.Dispose() }
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
        $XmlPath     = Join-Path $ElkFile.Directory.FullName "$ProjectName.edb\ProjektInfo.xml"

        if (-not (Test-Path $XmlPath)) {
            Write-LogMessage "ProjektInfo.xml nicht gefunden für '$ProjectName': $XmlPath" "WARN"
            $Results.Add(@{ ProjectName = $ProjectName; Type = "Projekt"; Path = $ElkFile.FullName; Info = $null; Error = "ProjektInfo.xml fehlt" })
            continue
        }

        try {
            $Info = Get-ProjectInfoFromXmlContent -XmlContent (Read-XmlFile $XmlPath) -ProjectName $ProjectName
            Write-LogMessage "Projekt '$ProjectName': $($Info.Emails.Count) E-Mail-Adresse(n)"
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
            $Doc = Read-XmlFromZip $ZwFile.FullName
            if (-not $Doc) {
                Write-LogMessage "ProjektInfo.xml nicht im Archiv '$ProjectName' gefunden" "WARN"
                $Results.Add(@{ ProjectName = $ProjectName; Type = "Archiv"; Path = $ZwFile.FullName; Info = $null; Error = "ProjektInfo.xml fehlt im Archiv" })
                continue
            }
            $Info = Get-ProjectInfoFromXmlContent -XmlContent $Doc -ProjectName $ProjectName
            Write-LogMessage "Archiv '$ProjectName': $($Info.Emails.Count) E-Mail-Adresse(n)"
            $Results.Add(@{ ProjectName = $ProjectName; Type = "Archiv"; Path = $ZwFile.FullName; Info = $Info; Error = "" })
        }
        catch {
            Write-LogMessage "Archiv '$($ZwFile.FullName)' nicht lesbar (evtl. kein ZIP-Format): $_" "WARN"
            $Results.Add(@{ ProjectName = $ProjectName; Type = "Archiv"; Path = $ZwFile.FullName; Info = $null; Error = $_.Exception.Message })
        }
    }

    Write-LogMessage "Gefunden: $($Results.Count) Projekte/Archive"
    return $Results
}

# Liste 1: Projekt -> E-Mail (eine Zeile pro Kombination)
function New-ProjectUserList {
    param($Projects)

    $Rows = foreach ($P in $Projects) {
        $ModDate = if ($P.Info -and $P.Info.ModificationDate) { $P.Info.ModificationDate.ToString('dd.MM.yyyy HH:mm') } else { "" }

        if ($P.Info -and $P.Info.Emails.Count -gt 0) {
            foreach ($Mail in $P.Info.Emails.Keys) {
                [PSCustomObject]@{
                    'Projektname'     = $P.ProjectName
                    'E-Mail'          = $Mail
                    'Fundstelle'      = ($P.Info.Emails[$Mail] -join ', ')
                    'Typ'             = $P.Type
                    'Letzte Änderung' = $ModDate
                    'Pfad'            = $P.Path
                    'Hinweis'         = $P.Error
                }
            }
        }
        elseif ($Config.IncludeProjectsWithoutEmail) {
            [PSCustomObject]@{
                'Projektname'     = $P.ProjectName
                'E-Mail'          = ""
                'Fundstelle'      = ""
                'Typ'             = $P.Type
                'Letzte Änderung' = $ModDate
                'Pfad'            = $P.Path
                'Hinweis'         = if ($P.Error) { $P.Error } else { "Keine E-Mail gefunden" }
            }
        }
    }

    return @($Rows | Sort-Object 'Projektname', 'E-Mail', 'Pfad')
}

# Liste 2: E-Mail -> Projekte (eine Zeile pro Benutzer)
function New-UserProjectList {
    param($ProjectRows)

    $ProjectRows |
        Where-Object { $_.'E-Mail' } |
        Group-Object 'E-Mail' |
        Sort-Object Name |
        ForEach-Object {
            $Names = @($_.Group | Select-Object -ExpandProperty 'Projektname' | Sort-Object -Unique)
            [PSCustomObject]@{
                'E-Mail'        = $_.Name
                'Anzahl'        = $Names.Count
                'Projekte'      = ($Names -join ', ')
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

    $WithMail = @($Projects | Where-Object { $_.Info -and $_.Info.Emails.Count -gt 0 }).Count
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
#   EPLAN_Projekte_Benutzer.csv  - Projekt | E-Mail | Fundstelle | ... (sortiert nach Projekt)
#   EPLAN_Benutzer_Projekte.csv  - E-Mail | Anzahl | Projekte (sortiert nach E-Mail)
#
# Die Spalte "Fundstelle" zeigt, in welcher Property der ProjektInfo.xml die
# Adresse stand (z. B. "Created by"). Damit lässt sich prüfen, ob die richtigen
# Benutzerfelder erfasst werden.
#
# Nur Lesezugriff, keine externen Abhängigkeiten (PowerShell 5.1+).
# ============================================================================
