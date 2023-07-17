#Author: James Romeo Gaspar

$LogFilePath = "C:\Temp\ADComputerMovement_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$LogMessageFormat = "{0:s} {1}"
$ErrorActionPreference = "Stop"
$ADComps = Get-Content C:\Temp\ADComputers.txt
$TargetOU = "<Replace with your target OU>"

function Log-Message {
    param([string]$Message)
    $FormattedMessage = $LogMessageFormat -f (Get-Date), $Message
    Add-Content -Path $LogFilePath -Value $FormattedMessage
}

function Get-SourceOU {
    param([string]$ComputerName)
    try {
        $DistinguishedName = (Get-ADComputer -Identity $ComputerName).DistinguishedName
        $SourceOU = ($DistinguishedName -split ',',2)[1]
    }
    catch {
        $SourceOU = $null
    }
    return $SourceOU
}

function Move-ADComputerToOU {
    param([string]$ComputerName, [string]$TargetOU)

    $SourceOU = Get-SourceOU -ComputerName $ComputerName
    if ($SourceOU -ne $null) {
        $Computer = Get-ADComputer -Identity $ComputerName
        Move-ADObject -Identity $Computer -TargetPath $TargetOU -Confirm:$false
        $Message = "Moved computer object '$ComputerName' from '$SourceOU' to '$TargetOU'"
        Log-Message -Message $Message
        return 1
    }
    else {
        $Message = "Could not find computer object '$ComputerName'"
        Log-Message -Message $Message
        return 0
    }
}

try {
    $NumMoved = 0
    foreach ($ComputerName in $ADComps) {
        Move-ADComputerToOU -ComputerName $ComputerName -TargetOU $TargetOU
        if ($SourceOU -ne $null) {
            $NumMoved++
        }
    }

    $Message = "Moved $NumMoved computer objects to OU '$TargetOU'"
    Log-Message -Message $Message
}
catch {
    $ErrorMessage = "An error occurred: $_"
    Log-Message -Message $ErrorMessage
    Write-Host -ForegroundColor Red $ErrorMessage
}

$ErrorActionPreference = "Continue"
