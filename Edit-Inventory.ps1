<#
.SYNOPSIS
Manages the current inventory
.DESCRIPTION
User must have the excel workbook closed before running this script. User can make edits to the
inventory by adding equipment, removing equipment, and record equipment from left employees
.EXAMPLE
.\Edit-Inventory.ps1
#>

#Requires -Version 7.0
#Requires -Modules ImportExcel

$workbookPath = "C:\Users\estebamx\pshell\csv_parser\testbook.xlsx"

function Add-Equipment {
<#
.SYNOPSIS
Adds a new device to active inventory
#>
    param(
        [string]$User,
        [string]$Model,
        [string]$ServiceTag,
        [string]$SLAssetTag,
        [string]$Hostname,
        [string]$Swarmhost,
        [string]$Notes,
        [PSCustomObject]$NewDevice,
        [string]$WorksheetName = "Active Inventory"
    )
    if (-not $NewDevice) {    
        $NewDevice = [PSCustomObject]@{
            "User" = $User
            "Updated" = Get-Date
            "Desktop/Laptop Model" = $Model
            "Service Tag" = $ServiceTag.ToUpper()
            "SL Asset Tag" = $SLAssetTag
            "Computer Name" = $Hostname
            "Swarmhost?" = $Swarmhost
            "Notes" = $Notes
        }
    }
    $xlpkg = Open-ExcelPackage -Path $workbookPath
    $data = Import-Excel -ExcelPackage $xlpkg -WorksheetName $WorksheetName
    $exportExcelSplat = @{
        ExcelPackage = $xlpkg
        WorksheetName = $WorksheetName
        TableStyle = "Light14"
        ClearSheet = $true
        AutoSize = $true
        PassThru = $true
    }
    $newXlpkg = $data + $NewDevice | Sort-Object -Property "User" | Export-Excel @exportExcelSplat
    $setExcelColumnSplat = @{
        ExcelPackage = $newXlpkg
        WorksheetName = $WorksheetName
        Column = 2
        NumberFormat = "Short Date"
    }
    Set-ExcelColumn @setExcelColumnSplat
    Close-ExcelPackage -ExcelPackage $newXlpkg
    Write-Host "  Equipment added to $File file!"
}

function Remove-Inventory {
<#
.SYNOPSIS
Removes a device from active inventory
#>
    param(
        [string]$ServiceTag
    )
    $xlpkg = Open-ExcelPackage -Path $workbookPath
    $data = Import-Excel -ExcelPackage $xlpkg -WorksheetName "Active Inventory"
    if ($data."Service Tag" -notcontains $ServiceTag) {
        Close-ExcelPackage -ExcelPackage $xlpkg -NoSave
        Write-Warning "  Could not locate equipment with the service tag: $($ServiceTag.ToUpper())"
        return
    }
    $exportExcelSplat = @{
        ExcelPackage = $xlpkg
        WorksheetName = "Active Inventory"
        TableStyle = "Light14"
        ClearSheet = $true
        AutoSize = $true
        PassThru = $true
    }
    $newXlpkg = $data | Where-Object { $_."Service Tag" -ne $ServiceTag } |
    Export-Excel @exportExcelSplat
    $setExcelColumnSplat = @{
        ExcelPackage = $newXlpkg
        WorksheetName = "Active Inventory"
        Column = 2
        NumberFormat = "Short Date"
    }
    Set-ExcelColumn @setExcelColumnSplat
    Close-ExcelPackage -ExcelPackage $newXlpkg
    Write-Host "  Equipment removed!"
}

function Move-Employee {
<#
.SYNOPSIS
Moves employee equipment from active inventory to left company 
#>
    param(
        [string]$User
    )
    $xlpkg = Open-ExcelPackage -Path $workbookPath
    $data = Import-Excel -ExcelPackage $xlpkg -WorksheetName "Active Inventory"
    if ($data."User" -notcontains $User) {
        Close-ExcelPackage -ExcelPackage $xlpkg -NoSave
        Write-Warning "  Could not find user: $User"
        return
    }
    $userEquipment = $data | Where-Object { $_."User" -eq $User }
    $exportExcelSplat = @{
        ExcelPackage = $xlpkg
        WorksheetName = "Active Inventory"
        TableStyle = "Light14"
        ClearSheet = $true
        AutoSize = $true
        PassThru = $true
    }
    $newXlpkg = $data | Where-Object { $_."User" -ne $User } | Export-Excel @exportExcelSplat
    $setExcelColumnSplat = @{
        ExcelPackage = $newXlpkg
        WorksheetName = "Active Inventory"
        Column = 2
        NumberFormat = "Short Date"
    }
    Set-ExcelColumn @setExcelColumnSplat
    Close-ExcelPackage -ExcelPackage $newXlpkg
    foreach ($device in $userEquipment) {
        Add-Equipment -NewDevice $device -WorksheetName "Left Company"
    }
    Write-Host "  User moved."
}

# main logic
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
