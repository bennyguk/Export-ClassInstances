    <#
    .SYNOPSIS
        Add one or more file attachments to an SCSMObject (WorkItem or ConfigItem)

    .DESCRIPTION
        Add one or more file attachments to an SCSMObject (WorkItem or ConfigItem)

    .PARAMETER <ParameterName>
        <Parameter Decription>
    
    .EXAMPLE
        #Files to Objects
        Add-SCSMFileAttachment -SMObject $SMObject -Path E:\File.csv

        #Attach to objects via pipline
        $SMObject | Add-SCSMFileAttachment -Path E:\File.csv

        #Attach Files via GUID
        Add-SCSMFileAttachment -ID "3df4b654-f230-bd4a-528a-7c0df4f5f23b" -Path E:\File.csv

        #Attach Files in Folder
        Add-SCSMFileAttachment -SMObject $SMObject -Path E:\Folder

        #Attache Multiple Files or Folders (not Recusive)
        Add-SCSMFileAttachment -SMObject $SMObject -Path E:\File.csv, E:\Folder, E:\File2.csv
    
    .NOTES

    .LINK
#>

[CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact='Low')]
param
(
    <#
    [parameter(
        Mandatory=$false,
        ValueFromPipeline=$false,
        ValueFromPipelibebyPropertyName=$false,
        Position=0,
        ParameterSetName='Set1',
        HelpMessage="Enter Managment Server Computer Name")]
        [Microsoft.EnterpriseManagement.Common.EnterpriseManagementObject]
    [ValidateSet("Value1","Value2")] 
        $Example,       
    #>
    
    [parameter(
        ParameterSetName="Object",
        ValueFromPipeline=$true,
        Mandatory=$true)]
        [Microsoft.EnterpriseManagement.Common.EnterpriseManagementObject]
        $SMObject,

    [parameter(
        ParameterSetName="GUID",
        ValueFromPipeline=$true,
        Mandatory=$true)]
        [System.Guid]
        $ID,

    [parameter( 
        Mandatory = $true )]
        [string[]]$Path,

    [parameter(
        HelpMessage="Enter Managment Server Computer Name")]
        [string]
        $ComputerName


)#end param
BEGIN 
{ 
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

} #End BEGIN

PROCESS
{
    if ($PSCmdlet.ParameterSetName -eq "GUID")
    {
        $SMObject = Get-SCSMObject -Id $ID
    }

    if ($SMObject.ClassName -match "workitem") {
        $RelationshipClassName = "System.WorkItemHasFileAttachment"
        $ProjectionName =        "System.WorkItem.Projection"
        $ObjID =                    $SMObject.ID
        $ActionLogClass =        Get-SCSMClass -Name System.WorkItem.TroubleTicket.ActionLog$
        $ActionLogRel =          Get-SCSMRelationshipClass -Name System.WorkItemHasActionLog$
    } else {
        $RelationshipClassName = "System.ConfigItemHasFileAttachment"
        $ProjectionName =        "System.ConfigItem.Projection"
        $ObjID =                    $SMObject.Id
    }

    $ManagementGroup = New-Object Microsoft.EnterpriseManagement.EnterpriseManagementGroup $SMDefaultComputer
    $FileAttachmentRel = Get-SCSMRelationshipClass -Name $RelationshipClassName
    $FileAttachmentClass = Get-SCSMClass -Name System.FileAttachment$
    

    #Get the all files
    $AllFiles = $Path |%{Get-ChildItem $_}

    #Check how many files were in the directory
    #Also check for any empty files?

    Foreach ( $FileObject in $AllFiles ){
        #Create a filestream 
        $FileMode = [System.IO.FileMode]::Open
        $fRead = New-Object System.IO.FileStream $FileObject.FullName, $FileMode

        #Create file object to be inserted
        $NewFileAttach = New-Object Microsoft.EnterpriseManagement.Common.CreatableEnterpriseManagementObject($ManagementGroup, $FileAttachmentClass)
        #Populate properties with info
        $SCSMGUID_Attachment = [Guid]::NewGuid().ToString()
        $NewFileAttach.Item($FileAttachmentClass, "Id").Value = $SCSMGUID_Attachment
        $NewFileAttach.Item($FileAttachmentClass, "DisplayName").Value = $FileObject.Name
        $NewFileAttach.Item($FileAttachmentClass, "Description").Value = $FileObject.Name
        $NewFileAttach.Item($FileAttachmentClass, "Extension").Value = $FileObject.Extension
        $NewFileAttach.Item($FileAttachmentClass, "Size").Value = $FileObject.Length
        $NewFileAttach.Item($FileAttachmentClass, "AddedDate").Value = [DateTime]::Now.ToUniversalTime()
        $NewFileAttach.Item($FileAttachmentClass, "Content").Value = $fRead

        #Init projection
        $ProjectionType = Get-SCSMTypeProjection -Name $ProjectionName
        $Projection = Get-SCSMObjectProjection -Projection $ProjectionType -Filter "ID -eq $ObjID"


        #Attach file object to Service Manager
        $Projection.__base.Add($NewFileAttach, $FileAttachmentRel.Target)
        $Projection.__base.Commit()

        if ($SMObject.ClassName -match "workitem") {
            $SCSMGUID_ActionLog = [Guid]::NewGuid().ToString()
            $MP = Get-SCSMManagementPack -Name "System.WorkItem.Library"
            $ActionType = "System.WorkItem.ActionLogEnum.FileAttached"
            $NewLog = New-Object Microsoft.EnterpriseManagement.Common.CreatableEnterpriseManagementObject($ManagementGroup, $ActionLogClass)

            $NewLog.Item( $ActionLogClass, "Id").Value = $SCSMGUID_ActionLog
            $NewLog.Item( $ActionLogClass, "DisplayName").Value = $SCSMGUID_ActionLog
            $NewLog.Item( $ActionLogClass, "ActionType").Value = $MP.GetEnumerations().GetItem($ActionType)
            $NewLog.Item( $ActionLogClass, "Title").Value = "Attached File"
            $NewLog.Item( $ActionLogClass, "EnteredBy").Value = "SYSTEM"
            $NewLog.Item( $ActionLogClass, "Description").Value = $FileObject.Name
            $NewLog.Item( $ActionLogClass, "EnteredDate").Value = (Get-Date).ToUniversalTime()

            #Insert comment to action log
            $Projection.__base.Add($NewLog, $ActionLogRel.Target)
            $Projection.__base.Commit()
        }

        #Cleanup
        $fRead.Close();
    }

}#end PROCESS

END
{  }#end END

