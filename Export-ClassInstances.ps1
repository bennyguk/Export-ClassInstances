<#
.SYNOPSIS
    A script to export class instance properties (work item or configuration item based classes), all related items properties and file attachments to CSV file.
.DESCRIPTION
    This script could be useful if you need to export class instances in bulk for archival purposes, or if you need to make changes to a custom class
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
[Parameter(mandatory)]
[string] $ClassName,
[Parameter(mandatory)]
[string] $FilePath,
[string] $FileName = "Export.csv",
[Switch] $IncludePendingDelete = $false,
[parameter(HelpMessage="Enter Managment Server Computer Name")]
[string] $ComputerName = "localhost"
)
# Using the Service Manager cmdlets because I could not replicate the functionality of Get-SCSMRelationship in SMLets easily.
$GetInstallDirectory = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\System Center\2010\Service Manager\Setup' -Name InstallDirectory
$SMPSModule = $GetInstallDirectory.InstallDirectory + "Powershell\System.Center.Service.Manager.psd1"

Import-Module $SMPSModule

 #Set SMDefaultComputer
    If ($ComputerName)  
    { 
        $SMDefaultComputer = $ComputerName 
    }
    if(!$SMDefaultComputer)
    {
        Write-Error '$SMDefaultComputer is null in the current session and no -ComputerName parameter was passed to this function. Please specify one or the other.' -ErrorAction Stop
        break
    }

# Get the class information from Service Manager
$Class = Get-SCSMClass –name $ClassName

# Check to see if the class exists
if($Class -eq $Null) {
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

# Create hastable for CSV columns and values
$csvColumns = @{}

# Get relationship types for the class we are working with, either configuration Item or Work Item
if($Class.GetBaseTypes().Name -contains "System.ConfigItem") {
    $x = Get-SCSMClass -Id 62f0be9f-ecea-e73c-f00d-3dd78a7422fc
    $classRelationships = (Get-SCSMRelationship -Source $x).Name
}
else {
    $x = Get-SCSMClass -Id f59821e2-0364-ed2c-19e3-752efbb1ece9
    $classRelationships = (Get-SCSMRelationship -Source $x).Name
}

# Add each relionship type to the csvColumns hashtable
foreach ($classRelationship in $classRelationships) {
            $csvColumns[$classRelationship] = ""
}  

# Remove the Service Manager module and import SMLets for the rest of the script
Remove-Module System.Center.Service.Manager
Import-Module SMLets

# Create Hashtable to store CSV column names and values
$csvColumns = @{}

# Get all instances of the class. Only display Active items unless the IncludePendingDelete switch is used.
if($IncludePendingDelete) {
    $classInstances = Get-SCSMObject -Class $Class
}
else {
    $classInstances = Get-SCSMObject -Class $Class | Where-Object {$_.objectstatus -match "Active”}
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
        & $FilePath\Get-FileAttachments.ps1 -Id $classInstance.Get_Id() -ArchiveRootPath $FilePath\ExportedAttachments -ComputerName $smdefaultcomputer
        }       
        foreach ($Property in $classProperties) {
                $csvColumns[$Property] = $classInstance.$Property
        }
        foreach ($relationship in $relationshipDetails) {
            if($csvColumns.ContainsKey((Get-SCSMRelationshipClass -Id $relationship.RelationshipId).Name)) {
            $existingValue = $csvColumns[(Get-SCSMRelationshipClass -Id $relationship.RelationshipId).Name]
            $csvColumns[(Get-SCSMRelationshipClass -Id $relationship.RelationshipId).Name] = $existingValue + "," + $relationship.TargetObject.DisplayName
            }
            else {
            $csvColumns[(Get-SCSMRelationshipClass -Id $relationship.RelationshipId).Name] = $relationship.TargetObject.DisplayName
            }
        }      
        $results = New-Object PSobject -Property $csvColumns
        $results | Export-csv $FilePath\$fileName -NoTypeInformation -Append
}

