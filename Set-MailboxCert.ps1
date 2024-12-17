<#

.SYNOPSIS
    Lädt Zertifikate aus Zertifikat-Dateien, extrahiert die E-Mail-Adresse oder den RFC822-Namen aus dem SAN-Feld und aktualisiert die Benutzerzertifikate im Exchange Server.

.DESCRIPTION
    Dieses Skript durchsucht alle Zertifikat-Dateien in einem angegebenen Ordner, lädt die Zertifikate, prüft auf eine SAN-Erweiterung (Subject Alternative Name) und extrahiert die E-Mail-Adresse oder den RFC822-Namen. Anschließend wird das Zertifikat der entsprechenden Mailbox im Exchange Server zugewiesen.

.NOTES
    Version:    2024-06-17
    Autor:      Julius Prinz

.PARAMETER CertPath
    Der Pfad zum Ordner, der die PEM-Zertifikatsdateien enthält.

.PARAMETER Extension
    Der DateiExtension für die Zertifikatsdateien.

#>


<# Parameter #>
param(
    [Parameter(Mandatory = $true, HelpMessage = "Geben Sie den Pfad zu den Zertifikatsdateien an.")]
    [string]$CertPath,

    [Parameter(Mandatory = $false, HelpMessage = "Geben Sie den DateiExtension für die Zertifikatsdateien an (z.B. *.pem).")]
    [string]$Extension = "pem"
)


<# Variablen #>
$ErrorActionPreference = 'Stop'


<# Programm #>
$Extension = "*.$Extension"
foreach ($file in Get-ChildItem -Path $CertPath -File -Filter $Extension -Recurse) {
    # Zertifikat aus Datei laden
    try {
        $cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($file.FullName)
    }
    catch {
        Write-Warning "Fehler beim Laden des Zertifikats: $($file.FullName). $_"
        continue
    }

    # SAN-Erweiterung extrahieren
    $sanExtension = $cert.Extensions | Where-Object { $_.Oid.Value -eq '2.5.29.17' }
    if ($sanExtension) {
        # Inhalt der SAN-Erweiterung formatieren
        $sanContent = $sanExtension.Format(0)

        # Suche nach Mail-Adresse oder RFC822-Name
        $sanMail = ($sanContent -split '\r?\n' | Where-Object { $_ -match "email:" -or $_ -match "RFC822-Name" }) `
            -replace '.*(?:email:|RFC822-Name=)([^,]+).*', '$1'

        if ($sanMail) {
            Write-Output "Verwende E-Mail/RFC822-Name aus SAN: $sanMail"

            # Set-Mailbox-Befehl ausführen
            try {
                Set-Mailbox -Identity $sanMail -UserCertificate (, $cert.GetRawCertData())
                Write-Output "Zertifikat erfolgreich zugewiesen für: $sanMail"
            }
            catch {
                Write-Warning "Fehler beim Zuweisen des Zertifikats für Mailbox $($sanMail); $_"
            }
        }
        else {
            Write-Warning "Keine E-Mail-Adresse oder RFC822-Name im SAN gefunden für Zertifikat: $($file.FullName)"
        }
    }
    else {
        Write-Warning "Keine SAN-Erweiterung gefunden für Zertifikat: $($file.FullName)"
    }
}


<# Signatur #>