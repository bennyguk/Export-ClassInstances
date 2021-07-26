<#
.SYNOPSIS
    A script to export System Center Service Manager class instance properties, relationships and file attachments of Work Item or Configuration Item based classes.
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
    [Switch] $IncludePendingDelete = $False,
    [parameter(Mandatory, HelpMessage = "Enter Managment Server Computer Name")]
    [string] $ComputerName
)

$GetInstallDirectory = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\System Center\2010\Service Manager\Setup' -Name InstallDirectory

$SMPSModule = $GetInstallDirectory.InstallDirectory + "Powershell\System.Center.Service.Manager.psd1"

Import-Module $SMPSModule

#Set the SMDefaultComputerparameter for SMlets
$SMDefaultComputer = $ComputerName 

# Get the class information from Service Manager
$Class = Get-SCSMClass -ComputerName $ComputerName | Where-Object { $_.Name -eq $ClassName }

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

function Get-FileAttachments {
    param 
    ([Guid] $Id)
    
    $WIhasAttachMent = "aa8c26dc-3a12-5f88-d9c7-753e5a8a55b4"
    $CIhasAttachMent = "095ebf2a-ee83-b956-7176-ab09eded6784"
 
    # Get Enterprise Management Object
    $Emo = Get-SCSMObject -Id $Id -ComputerName $ComputerName
 
    # Check if this is a Work Item or a Configuration Item to make sure we use the correct relationship
    $WIhasAttachMentClass = Get-SCSMRelationshipClass -Id $WIhasAttachMent -ComputerName $ComputerName
    $WIClass = Get-SCSMClass System.WorkItem$ -ComputerName $ComputerName

    if ($Emo.IsInstanceOf($WIClass)) {
        $files = Get-SCSMRelatedObject -SMObject $Emo -Relationship $WIhasAttachMentClass -ComputerName $ComputerName
    }
    else {
        $CIhasAttachMentClass = Get-SCSMRelationshipClass -Id $CIhasAttachMent -ComputerName $ComputerName
        $CIClass = Get-SCSMClass System.ConfigItem$ -ComputerName $ComputerName
        if ($Emo.IsInstanceOf($CIClass)) {
            $files = Get-SCSMRelatedObject -SMObject $Emo -Relationship $CIhasAttachMentClass -ComputerName $ComputerName
        }
        else {
            Write-Error "The Class type $Class is not supported" -ErrorAction Stop
        }
    }
 
    # For each file, archive to folder named with the ID of the class instance
    if ($files) {
        $nArchivePath = $FilePath + "\ExportedAttachments\" + $Emo.Id
        New-Item -Path ($nArchivePath) -ItemType "directory" -Force | Out-Null
 
        foreach ($file in $files) {
            $fileCounter++
            Write-Progress -Id 1 -Status "Processing $($FileCounter) of $($files.count)" -Activity "Exporting files" -CurrentOperation $file.DisplayName -PercentComplete (($fileCounter / $files.count) * 100)
            Try {
                $fs = [IO.File]::OpenWrite(($nArchivePath + "\" + $file.DisplayName))
                $memoryStream = New-Object IO.MemoryStream
                $buffer = New-Object byte[] 8192
                [int]$bytesRead | Out-Null
                while (($bytesRead = $file.Content.Read($buffer, 0, $buffer.Length)) -gt 0) {
                    $memoryStream.Write($buffer, 0, $bytesRead)
                }        
                $memoryStream.WriteTo($fs)
            }
            Finally {
                $fs.Close()
                $memoryStream.Close()
            }
        }
    }
}

# Create Hashtables to store CSV column names and values
$csvColumns = @{}
$csvRelColumns = @{}

# Get relationship types for the class we are working with
foreach ($baseType in $class.GetBaseTypes()) {
    $classRelationships = (Get-SCSMRelationship -ComputerName $ComputerName -Source $baseType).Name
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

# Collect the property names for the class and add to an array.
$classProperties = @()
foreach ($property in @($Class) + @($Class.GetBaseTypes())) {
    $classProperties += $property.PropertyCollection
}
$classProperties = $classProperties.Name

# Get class instance and related item property values and output to a CSV file. Export any attachments to a subdirectory called ExportedAttachments
foreach ($classInstance in $classInstances) {
    $counter++
    Write-Progress -Id 0 -Status "Processing $($counter) of $($classInstances.count)" -Activity "Exporting all instances of $($class.DisplayName)" -CurrentOperation $classInstance.DisplayName -PercentComplete (($counter / $classInstances.count) * 100)
    $relationshipDetails = Get-SCSMRelationshipObject -BySource $classInstance
    Get-FileAttachments -Id $classInstance.Get_Id()

    foreach ($Property in $classProperties) {
        $csvColumns[$Property] = $classInstance.$Property
    }

    # If the class does not have a key property, add the internal ID instead
    if (!$class.GetKeyProperties()) {
        $csvColumns["ID"] = $classInstance.ID
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
    # Add the key property of the class instance to the relationships CSV file as a reference. If there is no key, add the internal ID
    if (!$class.GetKeyProperties()) {
        $csvRelColumns["ID"] = $classInstance.ID
    }
    else {
        $csvRelColumns[$class.GetKeyProperties().Name] = $classInstance.ID
    }
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
