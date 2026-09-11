<#
.Synopsis
Manages the current inventory
.Description
User edits two csv files by adding equipment to the active inventory, removing equipment from the
active inventory, and transfer previously loaned equpment of left employees for record
.Parameter ActiveInventoryFile
A csv file from the current directory representing the active inventory
.Parameter LeftCompanyFile
A csv file from the current directory representing employees who left and their old equipment
.Example
.\Edit-Inventory.ps1 active_inventory.csv left_company.csv
#>

#requires -version 7.0

param(
    [Parameter(
        Mandatory,
        Position = 0,
        HelpMessage = "A csv file from the current directory representing the active inventory"
    )]
    [string]$ActiveInventoryFile,

    [Parameter(
        Mandatory,
        Position = 1,
        HelpMessage = ("A csv file from the current directory representing employees who left " +
            "and their old equipment")
    )]
    [string]$LeftCompanyFile
)

try {
    [void](Import-Csv -Path "$(Get-Location)\$ActiveInventoryFile" -ErrorAction Stop)
} catch {
    throw "Failed to find active inventory file in the current directory.`n$($_.Exception.Message)"
}

try {
    [void](Import-Csv -Path "$(Get-Location)\$LeftCompanyFile" -ErrorAction Stop)
} catch {
    throw "Failed to find left company file in the current directory.`n$($_.Exception.Message)"
}

function Add-Equipment {
    param(
        [string]$User,
        [string]$Model,
        [string]$ServiceTag,
        [string]$SLAssetTag,
        [string]$Hostname,
        [string]$Swarmhost,
        [string]$Notes,
        [PSCustomObject]$NewDevice,
        [string]$File = $ActiveInventoryFile
    )
    if (-not $NewDevice) {    
        $NewDevice = [PSCustomObject]@{
            "User" = $User
            "Updated" = (Get-Date).toShortDateString()
            "Desktop/Laptop Model" = $Model
            "Service Tag" = $ServiceTag.ToUpper()
            "SL Asset Tag" = $SLAssetTag
            "Computer Name" = $Hostname
            "Swarmhost?" = $Swarmhost
            "Notes" = $Notes
        }
    }
    (Import-Csv -Path "$(Get-Location)\$File") + $NewDevice |
    Sort-Object -Property "User" |
    Export-Csv -Path "$(Get-Location)\$File" -NoTypeInformation
    Write-Host "  Equipment added to $File file!"
}

function Remove-Inventory {
    param(
        [string]$ServiceTag
    )
    $activeInventory = (Import-Csv -Path "$(Get-Location)\$ActiveInventoryFile")
    if ($activeInventory."Service Tag" -notcontains $ServiceTag) {
        Write-Warning "  Could not locate equipment with the service tag: $($ServiceTag.ToUpper())"
        return
    }
    $activeInventory | Where-Object { $_."Service Tag" -ne $ServiceTag } |
    Export-Csv -Path "$(Get-Location)\$ActiveInventoryFile" -NoTypeInformation
    Write-Host "  Equipment removed!"
}

function Move-Employee {
    param(
        [string]$User
    )
    $activeInventory = (Import-Csv -Path "$(Get-Location)\$ActiveInventoryFile")
    if ($activeInventory."User" -notcontains $User) {
        Write-Warning "  Could not find user: $User"
        return
    }
    $userEquipment = $activeInventory | Where-Object { $_."User" -eq $User }
    $activeInventory | Where-Object { $_."User" -ne $User } |
    Export-Csv -Path "$(Get-Location)\$ActiveInventoryFile" -NoTypeInformation
    foreach ($device in $userEquipment) {
        device."Updated" = (Get-Date).toShortDateString()
        Add-Equipment -NewDevice $device -File $LeftCompanyFile
    }
    Write-Host "  User moved."
}

Write-Host "Hello, this is a powershell script that can help you with your current inventory."
$response = Read-Host ("Would you like to (1) add, (2) remove, or (3) transfer employee equipment? " +
    "To end script type (N). (1/2/3/N)")
Write-Host
while ($response -ne "N") {
    if ($response -eq "1") {
        $equipmentParams = @{
            User = Read-Host "  Who does this device belong to?"
            Model = Read-Host "  What is the model of this device?"
            ServiceTag = Read-Host "  What is the service tag of this device?"
            SLAssetTag = Read-Host "  What is the SL asset tag of this device?"
            Hostname = Read-Host "  What is the computer name of this device?"
            Swarmhost = Read-Host "  Is this computer a swarmhost?"
            Notes = Read-Host "  Any additional notes you want to add"
        }
        Add-Equipment @equipmentParams
        Write-Host
    } elseif ($response -eq "2") {
        $serviceTag = Read-Host "  Type the service tag of the equipment you want to remove"
        Remove-Inventory -ServiceTag $serviceTag
        Write-Host
    } elseif ($response -eq "3") {
        $user = Read-Host "  Type the full name of the user you want transfer"
        Move-Employee -User $user
        Write-Host
    } else {
        Write-Warning "  Invalid choice"
        Write-Host
    }
    $response = Read-Host ("Would you like to (1) add, (2) remove, or (3) transfer employee " + 
        "equipment? To end script type (N). (1/2/3/N)")
    Write-Host
}
