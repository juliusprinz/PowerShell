<#

.SYNOPSIS
    Erstellt eine neue virtuelle Maschine mit den angegebenen Parametern.

.DESCRIPTION
    Dieses Skript erstellt eine neue virtuelle Maschine und fragt interaktiv den Namen, die RAM-Größe und die VM-Version ab.

.NOTES
    Version:    2024-12-16
    Autor:      Julius Prinz

#>


<# Parameter #>
param()


<# Variablen #>

# Benutzerinteraktion: Abfrage des VM-Namens
$VMName = Read-Host "Bitte geben Sie den Namen der virtuellen Maschine ein"

# Benutzerinteraktion: Abfrage der RAM-Größe in GB
$RAMSizeGB = Read-Host "Bitte geben Sie die RAM-Größe in GB ein"
$MemoryStartupBytes = [int]$RAMSizeGB * 1GB

# Benutzerinteraktion: Abfrage der gewünschten VM-Version
Write-Host "Server 2019: 9.0"
Write-Host "Server 2022: 11.0"
Write-Host "Server 2025: 12.0"
$VMVersion = Read-Host "Bitte geben Sie die gewünschte VM-Version ein (z. B. 9.0)"

<# Funktionen #>


<# Programm #>

# Neue virtuelle Maschine erstellen
New-VM -Name $VMName -MemoryStartupBytes $MemoryStartupBytes -Generation 2 -Version $VMVersion

Write-Host "Die virtuelle Maschine '$VMName' mit der Version $VMVersion wurde erfolgreich erstellt."


<# Signatur #>