# Export-ClassInstances
A script to export class instance properties, relationships and file attachment from System Center Service Manager (Work Item or Configuration item based classes).
    
This script could be usful if you need to export class instances in bulk for archival purposes, or if you need to make changes to a custom class that are not upgrade compatible for later import with Import-ClassInstances.ps1 (https://github.com/bennyguk/Import-ClassInstances).



## To use
The script requires SMLets to be installed (https://github.com/SMLets/SMLets/releases) as well as the cmdlets distributed with the Service Manager console.

The script has the following parameters:
* ClassName - The Name property of the class rather than DisplayName to be exported.
* FilePath - A folder path to save the CSV files and file attachments.
* ComputerName - The hostname of your SCSM management server.
* IncludePendingDelete - (Optional) If set to true, the script will export class instances that are pending to be deleted. Defaults to false.
* FileName - (Optional) - A name for the CSV file that will store the property information for each class instance. A separate CSV file will also be created for relationship information that appends the value of FileName with *-relationships.csv*. Defaults to Export.csv

## Additional Information
This script started life simply as a way of exporting all instances of a particular class from our production environment in preparation of an upgrade to a custom Configuration Item based class Management Pack, but i soon started to wonder if it would be possible to programmatically export properties and relationships of any Work Item or Configuration Item. 

It is very much a work in progress and does have limitations with some relationship types such as SLAs, Request Offerings and Billable Time user relationships. If you have an idea about how to deal with these relationship types or any other improvements, I would be delighted for you to contribute :)
