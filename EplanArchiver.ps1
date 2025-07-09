# ============================================================================

# EPLAN Projekt Archivierungstool

# ============================================================================

# Automatisches Finden und Verschieben alter EPLAN-Projekte

# Erstellt von: PowerShell Script

# Version: 1.0

# ============================================================================

#Requires -Version 5.1

# Globale Variablen und Konfiguration

# ============================================================================

# KONFIGURATION - Hier können Sie die Einstellungen anpassen

$Config = @{
# Quellverzeichnis für die Suche (hier anpassen!)
SourcePath = “C:\EPLAN_Projekte”

```
# Zielverzeichnis für alte Projekte (hier anpassen!)
ArchivePath = "C:\EPLAN_Archiv"

# Anzahl Monate für "alte" Projekte
MonthsOld = 6

# Dry-Run Modus (True = nur testen, False = tatsächlich verschieben)
DryRun = $true

# CSV-Datei für Berichte
CsvReportPath = "C:\Temp\EPLAN_Archiv_Report.csv"

# Log-Datei
LogPath = "C:\Temp\EPLAN_Archiver.log"
```

}

# Globale Variablen

$Global:FoundProjects = @()
$Global:ProcessedProjects = @()

# ============================================================================

# HILFSFUNKTIONEN

# ============================================================================

# Funktion: Schreibt Nachrichten in Log-Datei und Konsole

function Write-LogMessage {
param(
[string]$Message,
[string]$Level = “INFO”
)

```
$TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$LogEntry = "[$TimeStamp] [$Level] $Message"

# In Konsole ausgeben
switch ($Level) {
    "ERROR" { Write-Host $LogEntry -ForegroundColor Red }
    "WARN"  { Write-Host $LogEntry -ForegroundColor Yellow }
    "INFO"  { Write-Host $LogEntry -ForegroundColor Green }
    default { Write-Host $LogEntry }
}

# In Log-Datei schreiben
try {
    Add-Content -Path $Config.LogPath -Value $LogEntry -Encoding UTF8
}
catch {
    Write-Host "Fehler beim Schreiben in Log-Datei: $_" -ForegroundColor Red
}
```

}

# Funktion: Überprüft und erstellt notwendige Verzeichnisse

function Initialize-Directories {
Write-LogMessage “Initialisiere Verzeichnisse…”

```
# Prüfe Quellverzeichnis
if (-not (Test-Path $Config.SourcePath)) {
    Write-LogMessage "FEHLER: Quellverzeichnis '$($Config.SourcePath)' existiert nicht!" "ERROR"
    return $false
}

# Erstelle Zielverzeichnis falls notwendig
if (-not (Test-Path $Config.ArchivePath)) {
    try {
        New-Item -Path $Config.ArchivePath -ItemType Directory -Force | Out-Null
        Write-LogMessage "Zielverzeichnis '$($Config.ArchivePath)' erstellt."
    }
    catch {
        Write-LogMessage "FEHLER: Kann Zielverzeichnis nicht erstellen: $_" "ERROR"
        return $false
    }
}

# Erstelle Temp-Verzeichnis für Log
$TempDir = Split-Path $Config.LogPath
if (-not (Test-Path $TempDir)) {
    New-Item -Path $TempDir -ItemType Directory -Force | Out-Null
}

return $true
```

}

# Funktion: Parst das Unix-Timestamp aus der XML

function ConvertFrom-UnixTimestamp {
param([string]$UnixTimestamp)

```
try {
    $UnixTime = [long]$UnixTimestamp
    $DateTime = [datetime]'1970-01-01 00:00:00'
    return $DateTime.AddSeconds($UnixTime)
}
catch {
    Write-LogMessage "Fehler beim Konvertieren des Timestamps '$UnixTimestamp': $_" "WARN"
    return $null
}
```

}

# Funktion: Extrahiert Informationen aus ProjektInfo.xml

function Get-ProjectInfoFromXml {
param(
[string]$XmlPath,
[string]$ProjectName
)

```
try {
    if (-not (Test-Path $XmlPath)) {
        Write-LogMessage "ProjektInfo.xml nicht gefunden: $XmlPath" "WARN"
        return $null
    }
    
    [xml]$XmlContent = Get-Content $XmlPath -Encoding UTF8
    
    # Suche nach Modification Date Property
    $ModificationProperty = $XmlContent.SelectSingleNode("//Property[@name='Modification date']")
    
    if ($ModificationProperty) {
        $UnixTimestamp = $ModificationProperty.InnerText
        $ModificationDate = ConvertFrom-UnixTimestamp $UnixTimestamp
        
        if ($ModificationDate) {
            Write-LogMessage "Projekt '$ProjectName': Letzte Änderung am $($ModificationDate.ToString('dd.MM.yyyy HH:mm:ss'))"
            
            return @{
                ProjectName = $ProjectName
                ModificationDate = $ModificationDate
                UnixTimestamp = $UnixTimestamp
                XmlPath = $XmlPath
            }
        }
    }
    else {
        Write-LogMessage "Keine Modification Date in XML gefunden für Projekt '$ProjectName'" "WARN"
    }
}
catch {
    Write-LogMessage "Fehler beim Lesen der XML '$XmlPath': $_" "ERROR"
}

return $null
```

}

# Funktion: Extrahiert ProjektInfo.xml aus .zw1 Archive (nur native PowerShell)

function Get-ProjectInfoFromArchive {
param(
[string]$ArchivePath,
[string]$ProjectName
)

```
try {
    Write-LogMessage "Versuche Archiv-Extraktion für: $ProjectName"
    
    # Temporäres Verzeichnis für Extraktion
    $TempExtractPath = Join-Path $env:TEMP "EPLAN_Extract_$([System.Guid]::NewGuid())"
    New-Item -Path $TempExtractPath -ItemType Directory -Force | Out-Null
    
    # .zw1 Dateien sind oft ZIP-Archive mit anderem Namen
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $ZipArchive = [System.IO.Compression.ZipFile]::OpenRead($ArchivePath)
    
    # Suche nach ProjektInfo.xml im Archiv
    $XmlEntry = $ZipArchive.Entries | Where-Object { $_.Name -eq "ProjektInfo.xml" }
    
    if ($XmlEntry) {
        $XmlPath = Join-Path $TempExtractPath "ProjektInfo.xml"
        [System.IO.Compression.ZipFileExtensions]::ExtractToFile($XmlEntry, $XmlPath, $true)
        $ZipArchive.Dispose()
        
        if (Test-Path $XmlPath) {
            $ProjectInfo = Get-ProjectInfoFromXml -XmlPath $XmlPath -ProjectName $ProjectName
            
            # Cleanup
            Remove-Item $TempExtractPath -Recurse -Force -ErrorAction SilentlyContinue
            
            return $ProjectInfo
        }
    }
    else {
        Write-LogMessage "ProjektInfo.xml nicht im Archiv gefunden: $ProjectName" "WARN"
        $ZipArchive.Dispose()
    }
    
    # Cleanup bei Fehler
    Remove-Item $TempExtractPath -Recurse -Force -ErrorAction SilentlyContinue
}
catch {
    Write-LogMessage "Archiv '$ArchivePath' konnte nicht verarbeitet werden (möglicherweise kein ZIP-Format): $_" "WARN"
}

return $null
```

}

# ============================================================================

# HAUPTFUNKTIONEN

# ============================================================================

# Funktion: Sucht nach EPLAN-Projekten im Verzeichnis

function Find-EplanProjects {
param([string]$SearchPath)

```
Write-LogMessage "Durchsuche Verzeichnis: $SearchPath"

$FoundProjects = @()
$CutoffDate = (Get-Date).AddMonths(-$Config.MonthsOld)

try {
    # Suche nach .elk Dateien (EPLAN Projekte)
    $ElkFiles = Get-ChildItem -Path $SearchPath -Filter "*.elk" -Recurse -ErrorAction SilentlyContinue
    
    foreach ($ElkFile in $ElkFiles) {
        $ProjectName = [System.IO.Path]::GetFileNameWithoutExtension($ElkFile.Name)
        $ProjectDir = Join-Path $ElkFile.Directory.FullName "$ProjectName.edb"
        
        Write-LogMessage "Gefundenes Projekt: $ProjectName"
        
        # Prüfe ob .edb Verzeichnis existiert
        if (Test-Path $ProjectDir) {
            # Suche nach ProjektInfo.xml im .edb Verzeichnis
            $XmlPath = Join-Path $ProjectDir "ProjektInfo.xml"
            $ProjectInfo = Get-ProjectInfoFromXml -XmlPath $XmlPath -ProjectName $ProjectName
            
            if ($ProjectInfo -and $ProjectInfo.ModificationDate -lt $CutoffDate) {
                $FoundProjects += @{
                    ProjectName = $ProjectName
                    ElkPath = $ElkFile.FullName
                    EdbPath = $ProjectDir
                    ModificationDate = $ProjectInfo.ModificationDate
                    Type = "Standard"
                    OriginalPath = $ElkFile.Directory.FullName
                }
                Write-LogMessage "Projekt '$ProjectName' ist älter als $($Config.MonthsOld) Monate und wird archiviert."
            }
            elseif ($ProjectInfo) {
                Write-LogMessage "Projekt '$ProjectName' ist noch aktuell (geändert: $($ProjectInfo.ModificationDate.ToString('dd.MM.yyyy')))"
            }
        }
        else {
            Write-LogMessage "Warnung: .edb Verzeichnis für '$ProjectName' nicht gefunden!" "WARN"
        }
    }
    
    # Suche nach .zw1 Dateien (EPLAN Archive)
    $ZwFiles = Get-ChildItem -Path $SearchPath -Filter "*.zw1" -Recurse -ErrorAction SilentlyContinue
    
    foreach ($ZwFile in $ZwFiles) {
        $ProjectName = [System.IO.Path]::GetFileNameWithoutExtension($ZwFile.Name)
        
        Write-LogMessage "Gefundenes Archiv: $ProjectName"
        
        $ProjectInfo = Get-ProjectInfoFromArchive -ArchivePath $ZwFile.FullName -ProjectName $ProjectName
        
        if ($ProjectInfo -and $ProjectInfo.ModificationDate -lt $CutoffDate) {
            $FoundProjects += @{
                ProjectName = $ProjectName
                ArchivePath = $ZwFile.FullName
                ModificationDate = $ProjectInfo.ModificationDate
                Type = "Archive"
                OriginalPath = $ZwFile.Directory.FullName
            }
            Write-LogMessage "Archiv '$ProjectName' ist älter als $($Config.MonthsOld) Monate und wird archiviert."
        }
        elseif ($ProjectInfo) {
            Write-LogMessage "Archiv '$ProjectName' ist noch aktuell (geändert: $($ProjectInfo.ModificationDate.ToString('dd.MM.yyyy')))"
        }
    }
    
}
catch {
    Write-LogMessage "Fehler beim Durchsuchen von '$SearchPath': $_" "ERROR"
}

Write-LogMessage "Gefunden: $($FoundProjects.Count) alte Projekte zum Archivieren"
return $FoundProjects
```

}

# Funktion: Verschiebt ein einzelnes Projekt

function Move-Project {
param(
[hashtable]$Project,
[bool]$DryRun = $true
)

```
$ProjectName = $Project.ProjectName
$TargetDir = Join-Path $Config.ArchivePath $ProjectName

try {
    if ($DryRun) {
        Write-LogMessage "[DRY-RUN] Würde Projekt '$ProjectName' nach '$TargetDir' verschieben"
        return @{
            Success = $true
            OriginalPath = $Project.OriginalPath
            NewPath = $TargetDir
            Action = "DRY-RUN"
        }
    }
    else {
        # Erstelle Zielverzeichnis
        if (-not (Test-Path $TargetDir)) {
            New-Item -Path $TargetDir -ItemType Directory -Force | Out-Null
        }
        
        if ($Project.Type -eq "Standard") {
            # Verschiebe .elk Datei und .edb Verzeichnis
            $TargetElk = Join-Path $TargetDir (Split-Path $Project.ElkPath -Leaf)
            $TargetEdb = Join-Path $TargetDir (Split-Path $Project.EdbPath -Leaf)
            
            Move-Item -Path $Project.ElkPath -Destination $TargetElk -Force
            Move-Item -Path $Project.EdbPath -Destination $TargetEdb -Force
            
            Write-LogMessage "Projekt '$ProjectName' erfolgreich nach '$TargetDir' verschoben"
        }
        else {
            # Verschiebe .zw1 Archiv
            $TargetArchive = Join-Path $TargetDir (Split-Path $Project.ArchivePath -Leaf)
            Move-Item -Path $Project.ArchivePath -Destination $TargetArchive -Force
            
            Write-LogMessage "Archiv '$ProjectName' erfolgreich nach '$TargetDir' verschoben"
        }
        
        return @{
            Success = $true
            OriginalPath = $Project.OriginalPath
            NewPath = $TargetDir
            Action = "MOVED"
        }
    }
}
catch {
    Write-LogMessage "FEHLER beim Verschieben von Projekt '$ProjectName': $_" "ERROR"
    return @{
        Success = $false
        OriginalPath = $Project.OriginalPath
        NewPath = $TargetDir
        Action = "ERROR"
        Error = $_.Exception.Message
    }
}
```

}

# Funktion: Erstellt CSV-Bericht

function New-CsvReport {
param([array]$ProcessedProjects)

```
Write-LogMessage "Erstelle CSV-Bericht..."

try {
    # CSV-Daten vorbereiten
    $CsvData = @()
    
    foreach ($Project in $ProcessedProjects) {
        $CsvData += [PSCustomObject]@{
            'Projektname' = $Project.ProjectName
            'Typ' = $Project.Type
            'Letzte Änderung' = $Project.ModificationDate.ToString('dd.MM.yyyy HH:mm:ss')
            'Ursprünglicher Pfad' = $Project.OriginalPath
            'Neuer Pfad' = if($Project.Result.Success) { $Project.Result.NewPath } else { "FEHLER" }
            'Aktion' = $Project.Result.Action
            'Status' = if($Project.Result.Success) { "Erfolgreich" } else { "Fehlgeschlagen" }
            'Fehler' = if($Project.Result.Error) { $Project.Result.Error } else { "" }
        }
    }
    
    # Erstelle CSV-Datei
    $CsvData | Export-Csv -Path $Config.CsvReportPath -Delimiter ';' -Encoding UTF8 -NoTypeInformation
    
    Write-LogMessage "CSV-Bericht erstellt: $($Config.CsvReportPath)"
    Write-LogMessage "Die CSV-Datei kann in Excel geöffnet werden."
    
}
catch {
    Write-LogMessage "Fehler beim Erstellen des CSV-Berichts: $_" "ERROR"
}
```

}

# ============================================================================

# HAUPTPROGRAMM

# ============================================================================

function Start-EplanArchiver {
Write-LogMessage “EPLAN Projekt Archivierungstool gestartet”
Write-LogMessage “Konfiguration:”
Write-LogMessage “  Quellpfad: $($Config.SourcePath)”
Write-LogMessage “  Zielpfad: $($Config.ArchivePath)”
Write-LogMessage “  Alter: $($Config.MonthsOld) Monate”
Write-LogMessage “  Dry-Run: $($Config.DryRun)”
Write-LogMessage “  CSV-Bericht: $($Config.CsvReportPath)”

```
# Initialisierung
if (-not (Initialize-Directories)) {
    Write-LogMessage "Initialisierung fehlgeschlagen. Abbruch." "ERROR"
    return
}

# Projekte suchen
$Global:FoundProjects = Find-EplanProjects -SearchPath $Config.SourcePath

if ($Global:FoundProjects.Count -eq 0) {
    Write-LogMessage "Keine alten Projekte gefunden."
    return
}

# Projekte verarbeiten
Write-LogMessage "Verarbeite $($Global:FoundProjects.Count) Projekte..."

foreach ($Project in $Global:FoundProjects) {
    $Result = Move-Project -Project $Project -DryRun $Config.DryRun
    
    $Global:ProcessedProjects += @{
        ProjectName = $Project.ProjectName
        Type = $Project.Type
        ModificationDate = $Project.ModificationDate
        OriginalPath = $Project.OriginalPath
        Result = $Result
    }
}

# Bericht erstellen
New-CsvReport -ProcessedProjects $Global:ProcessedProjects

# Zusammenfassung
$SuccessCount = ($Global:ProcessedProjects | Where-Object { $_.Result.Success }).Count
$ErrorCount = ($Global:ProcessedProjects | Where-Object { -not $_.Result.Success }).Count

Write-LogMessage "ZUSAMMENFASSUNG:"
Write-LogMessage "  Verarbeitet: $($Global:ProcessedProjects.Count) Projekte"
Write-LogMessage "  Erfolgreich: $SuccessCount"
Write-LogMessage "  Fehler: $ErrorCount"

if ($Config.DryRun) {
    Write-LogMessage "Dies war ein DRY-RUN. Keine Dateien wurden verschoben."
    Write-LogMessage "Setzen Sie `$Config.DryRun = `$false für den echten Lauf."
}

Write-LogMessage "EPLAN Projekt Archivierungstool beendet"
```

}

# ============================================================================

# PROGRAMMSTART

# ============================================================================

# Starte das Programm

Start-EplanArchiver

# ============================================================================

# VERWENDUNG:

# ============================================================================

# 1. Passen Sie die Konfiguration am Anfang des Skripts an

# 2. Führen Sie das Skript mit DryRun = $true aus zum Testen

# 3. Prüfen Sie die Log-Datei und den CSV-Bericht

# 4. Setzen Sie DryRun = $false für den echten Lauf

# 5. Das Skript kann dann automatisch per Aufgabenplanung ausgeführt werden

# 

# KEINE EXTERNEN ABHÄNGIGKEITEN:

# - Läuft mit Standard PowerShell 5.1+

# - Verwendet nur native ZIP-Unterstützung

# - Erstellt CSV-Berichte (öffnen sich in Excel)

# ============================================================================