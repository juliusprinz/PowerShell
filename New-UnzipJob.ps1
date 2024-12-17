<#

.SYNOPSIS
    Entpackt alle .zip-Dateien in einem Verzeichnis und dessen Unterverzeichnissen mithilfe von 7-Zip
    und protokolliert den Vorgang in einer Logdatei.

.DESCRIPTION
    Dieses Script durchsucht ein angegebenes Verzeichnis und alle Unterverzeichnisse nach .zip-Dateien.
    Jede gefundene .zip-Datei wird mit 7-Zip in das gleiche Verzeichnis entpackt. Erfolgreiche oder fehlerhafte
    Entpackungen werden in einer Logdatei protokolliert.

.NOTES
    Version:    2024-12-14
    Autor:      Julius Prinz

.PARAMETER RootPath
    Das Verzeichnis, das nach .zip-Dateien durchsucht werden soll.

.PARAMETER SevenZipPath
    Der vollständige Pfad zur ausführbaren 7z.exe-Datei.

#>

<# Parameter #>
param(
    [Parameter(Mandatory = $true)]
    [string]$RootPath,

    [Parameter(Mandatory = $false)]
[string]$SevenZipPath = "$($env:ProgramFiles)\7-Zip\7z.exe"
)

<# Variablen #>
# Überprüfen, ob der Pfad zu 7z.exe korrekt ist
if (-not (Test-Path -Path $SevenZipPath)) {
    Write-Error "Die angegebene 7z.exe konnte nicht gefunden werden: $SevenZipPath"
    return
}

# Überprüfen, ob das Quellverzeichnis existiert
if (-not (Test-Path -Path $RootPath)) {
    Write-Error "Das angegebene Verzeichnis existiert nicht: $RootPath"
    return
}

# Erstellen des Log-Dateinamens
$LogFileName = (Get-Date -Format "yyMMdd-HHmmss") + ".log"
$LogFilePath = Join-Path -Path $RootPath -ChildPath $LogFileName

<# Funktionen #>
function Write-Log {
    param (
        [string]$Message
    )
    # Nachricht in die Logdatei schreiben
    Add-Content -Path $LogFilePath -Value $Message
}

function Extract-ZipFile {
    param (
        [string]$ZipFilePath,
        [string]$SevenZipExecutable
    )

    # Zielverzeichnis ist das gleiche wie das der Zip-Datei
    $TargetDirectory = [System.IO.Path]::GetDirectoryName($ZipFilePath)

    # Entpacken der Datei mit 7z.exe
    $process = Start-Process -FilePath $SevenZipExecutable -ArgumentList "x", "`"$ZipFilePath`"", "-o`"$TargetDirectory`"", "-y" -NoNewWindow -PassThru -Wait

    # Fehlerbehandlung
    if ($process.ExitCode -ne 0) {
        $LogMessage = "FEHLER: $ZipFilePath"
        Write-Warning "Fehler beim Entpacken der Datei: $ZipFilePath"
    } else {
        $LogMessage = "ERFOLG: $ZipFilePath"
        Write-Output "Erfolgreich entpackt: $ZipFilePath"
    }

    # Lognachricht schreiben
    Write-Log -Message $LogMessage
}

<# Programm #>
# Logdatei initialisieren
Write-Log -Message "Entpackung gestartet: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"

# Alle .zip-Dateien im Quellverzeichnis und seinen Unterverzeichnissen finden
$ZipFiles = Get-ChildItem -Path $RootPath -Recurse -Filter "*.zip" -File

if ($ZipFiles.Count -eq 0) {
    $NoFilesMessage = "Keine .zip-Dateien im Verzeichnis gefunden: $RootPath"
    Write-Log -Message $NoFilesMessage
    Write-Output $NoFilesMessage
    return
}

# Jede gefundene .zip-Datei entpacken
foreach ($ZipFile in $ZipFiles){
    Write-Output "Entpacken von: $($ZipFile.FullName)"
    Extract-ZipFile -ZipFilePath $ZipFile.FullName -SevenZipExecutable $SevenZipPath
}

Write-Log -Message "Entpackung beendet: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Output "Alle .zip-Dateien wurden verarbeitet. Logdatei: $LogFilePath"


<# Signatur #>