<#
.SYNOPSIS
    A script to export class instance properties (work item or configuration item based classes), all relationships and file attachments to CSV file.
.DESCRIPTION
    This script could be usful if you need to export class instances in bulk for archival purposes, or if you need to make changes to a custom class
    that are not upgrade compatible for later import with Import-ClassInstances.ps1 (https://github.com/bennyguk/Import-ClassInstances).
    
    For more information, please see https://github.com/bennyguk/Export-ClassInstances

.PARAMETER ClassName
    Specifies the class name you wish to work with.

.PARAMETER FilePath
    Specifies the path to the folder you wish to export file attachments and CSV file to.

.PARAMETER FileName
    Specifies name of the CSV file - Will default to Export.csv

.PARAMETER ComputerName
    Specifies the SCSM server to connect to.

.PARAMETER IncludePendingDelete
    Will include class instances that have been deleted in the export.
    
.EXAMPLE
    Export-ClassInstances.ps1 -ClassName MyClass -FilePath c:\MyClassExport -FileName MyClassExport.csv -ComputerName MySCSMServer -IncludePendingDelete
#>
Param (
    [parameter(Mandatory)][string] $ClassName,
    [parameter(Mandatory, HelpMessage = "Enter a path to the exported CSV directory, excluding the filename")][string] $FilePath,
    [string] $FileName = "Export.csv",
    [Switch] $IncludePendingDelete = $false,
    [parameter(HelpMessage = "Enter Managment Server Computer Name")]
    [string] $ComputerName = "localhost"
)

$GetInstallDirectory = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\System Center\2010\Service Manager\Setup' -Name InstallDirectory

$SMPSModule = $GetInstallDirectory.InstallDirectory + "Powershell\System.Center.Service.Manager.psd1"

Import-Module $SMPSModule

#Set SMDefaultComputer
If ($ComputerName) { 
    $SMDefaultComputer = $ComputerName 
}
if (!$SMDefaultComputer) {
    Write-Error '$SMDefaultComputer is null in the current session and no -ComputerName parameter was passed to this function. Please specify one or the other.' -ErrorAction Stop
    break
}

# Get the class information from Service Manager
$Class = Get-SCSMClass | Where-Object { $_.Name -eq $ClassName }

# Check to see if the class exists
if (!$Class) {
    Write-Host "Could not load class '$className'. Please check the class name and try again."
    Exit
}

# Check to see if the file path exists
If (!(Test-Path $FilePath)) {
    Write-Host "Could not find '$FilePath'. Please check the path name and try again."
    Exit
}
# Create an ExportedAttachements directory if it does not exist
if (!(Test-Path $FilePath\ExportedAttachments)) {
    New-Item -Path $FilePath -Name "ExportedAttachments" -ItemType "directory"
}

# Create Hashtables to store CSV column names and values
$csvColumns = @{}
$csvRelColumns = @{}

# Get relationship types for the class we are working with
foreach ($baseType in $class.GetBaseTypes()) {
    $classRelationships = (Get-SCSMRelationship -Source $baseType).Name
    # Add each relionship type to the csvRelColumns hashtable
    foreach ($classRelationship in $classRelationships) {
        $csvRelColumns[$classRelationship] = ""
    }  
}

# Remove the Service Manager module and import SMLets
Remove-Module System.Center.Service.Manager
Import-Module SMLets

# Get all instances of the class. Only display Active items unless the IncludePendingDelete parameter is used.
# Config Items use the ObjectStatus property and Work Items use the Status property.
if ($IncludePendingDelete) {
    $classInstances = Get-SCSMObject -Class $Class
}
else {
    $classInstances = Get-SCSMObject -Class $Class | Where-Object { $_.objectstatus -match "Active" }
    if (!$classInstances) {
        $classInstances = Get-SCSMObject -Class $Class | Where-Object { $_.status -match "Active" -or "Closed" -or "Resolved" }
    }
}

# Collect the property names for the class and add to an array
$classProperties = @()
foreach ($property in @($Class) + @($Class.GetBaseTypes())) {
    $classProperties += $property.PropertyCollection
}
$classProperties = $classProperties.Name

# Get class instance and related item property values and output to a CSV file. Export any attachments to a subdirectory called ExportedAttachments
foreach ($classInstance in $classInstances) {
    $relationshipDetails = Get-SCSMRelationshipObject -BySource $classInstance
    Try {
        & $filePath\Get-FileAttachments.ps1 -Id $classInstance.Get_Id() -ArchiveRootPath $FilePath\ExportedAttachments -ComputerName $smdefaultcomputer
    }
    Catch [System.Management.Automation.CommandNotFoundException] {
        Write-Host -ForegroundColor Red "ERROR: The Get-FileAttachments script could not be found. Please check the script file is in the same location as the Export-ClassInstances script and try again." 
    }
    foreach ($Property in $classProperties) {
        $csvColumns[$Property] = $classInstance.$Property
    }
    foreach ($relationship in $relationshipDetails) {
        $existingValue = ""
        if ($relationship.TargetObject.DisplayName) {
            $existingValue = $csvRelColumns[(Get-SCSMRelationshipClass -Id $relationship.RelationshipId).Name]
            if ($existingValue) {
                $csvRelColumns[(Get-SCSMRelationshipClass -Id $relationship.RelationshipId).Name] = $existingValue + "," + $relationship.TargetObject.DisplayName
            }
            else {
                $csvRelColumns[(Get-SCSMRelationshipClass -Id $relationship.RelationshipId).Name] = $relationship.TargetObject.DisplayName
            }    
        }      
    }
    # Add the key property of the class instance to the relationships CSV file as a reference.
    $csvRelColumns[($classInstance.GetProperties() | Where-Object { $_.Key -eq "True" }).Name] = $classInstance.(($classInstance.GetProperties() | Where-Object { $_.Key -eq "True" }).Name)
    # Write the results to the CSV files
    $outputCSV = New-Object PSobject -Property $csvColumns
    $outputCSV | Export-csv $FilePath\$fileName -NoTypeInformation -Append
    $outputRelCSV = New-Object PSobject -Property $csvRelColumns
    if ($FileName -match ".") {
        $splitName = $FileName.Split('.')
        $relFileName = "$($splitName[0])-relationships.$($splitName[1])"
    }
    else {
        $relFileName = "$FileName-relationships"
    }
    $outputRelCSV | Export-csv $FilePath\$relFileName -NoTypeInformation -Append

    # Clear the $csvColumns hastable
    $csvColumns2 = $csvColumns.Clone();
    foreach ($key in $csvColumns2.keys) { $csvColumns[$key] = ''; }
    $csvColumns2 = $null;

    # Clear the $csvRelColumns hastable
    $csvRelColumns2 = $csvRelColumns.Clone();
    foreach ($key in $csvRelColumns2.keys) { $csvRelColumns[$key] = ''; }
    $csvRelColumns2 = $null;
}
