<#
  	.SYNOPSIS
  	This script takes "Name" input from the pipeline and queries an Exchange server for mailbox size, broken down into total size, inbox size, sent items size and deleted items size. 
  	.DESCRIPTION
 	Get-MailboxSize report is intended to be used to periodically report on the size of mailboxes. This is designed to be run from a work station that has either the 
 	Exchange admin tools installed or implicitly remoted via PSRemoting. Instructions on that are included HERE <Add Link>. 
 	This script takes the parameter NAME either from the pipeline or as a positional parameter. The output is in the form of a custom object with the properties of 'Mailbox','Inbox (MB)','Sent Items (MB)',
 	'Deleted Items (MB)',and 'Total Size (MB)'

	.PARAMETER

  	.EXAMPLE
        Get-mailboxsizes Blinky
    retrieves the size info for user Blinky and displays to the screen. 

  	.EXAMPLE
        $Names = get-mailbox 
        Get-mailboxsizes -Name $Names -Verbose | export-csv "$env:userprofile\desktop\Powershell Reports\MailboxSizes.csv" -NoTypeInformation
    Line one retrievees a list of mailboxes from the Exchange server and line two calls the function passing that parameter and outputing to a CSV file on the users's desktop.
	
	.INPUTS
		System.String,System.Int32

	.OUTPUTS
		System.String

	.NOTES
		For more information about advanced functions, call Get-Help with any
		of the topics in the links listed below.

#>

Function Get-MailboxSizes{
    [cmdletbinding()]
     Param(
        [Parameter(Mandatory=$True,
                   ValueFromPipeline=$True,
                   HelpMessage='One or more Mailboxes')]
        [string[]]$Name)
    Begin{
        $samail = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://samail10.swiff-train.com/PowerShell/ -Authentication Kerberos  -Credential $creds
    Import-PSSession $samail
    }

    Process {
        ForEach ($Mailbox in $Name){
            Write-Verbose "Calculating the Total Size for $Mailbox"
            $TotalSize =[Math]::Round(((Get-MailboxStatistics -Identity $MailBox).totalitemsize -replace '(.*\()|,| [a-z]*\)', '')/1MB,2)
            
            Write-Verbose "Calculating the Inbox Size for $Mailbox"
            $InboxSize =[Math]::Round(((Get-MailboxFolderStatistics -Identity $MailBox -FolderScope Inbox | 
                Where-Object {$_.FolderPath -eq '/Inbox'}).FolderAndSubfolderSize -replace '(.*\()|,| [a-z]*\)', '')/1MB,2)
            
            Write-Verbose "Calculating the Sent Items Size for $Mailbox"
            $SentSize =[Math]::Round(((Get-MailboxFolderStatistics $MailBox -FolderScope SentItems |
                Where-Object {$_.FolderPath -eq '/Sent Items'}).FolderAndSubfolderSize -replace '(.*\()|,| [a-z]*\)', '')/1MB,2)
            
            Write-Verbose "Calculating the Deleted Items Size for $Mailbox"
            $DeletedSize =[Math]::Round(((Get-MailboxFolderStatistics $MailBox -FolderScope DeletedItems |
                Where-Object {$_.FolderPath -eq '/Deleted Items'}).FolderAndSubfolderSize -replace '(.*\()|,| [a-z]*\)', '')/1MB,2)

                        # construct output object and output to pipeline
            $properties = @{'Mailbox'           = $Mailbox;
                            'Inbox (MB)'        = $InboxSize;
                            'Sent Items (MB)'   = $SentSize;
                            'Deleted Items (MB)'= $DeletedSize;
                            'Total Size (MB)'   = $TotalSize;
                            }
            $MailObj = New-Object -TypeName PSObject -Property $properties
            $MailObj.psobject.typenames.insert(0,'ST.MailBox')
            Write-Output $MailObj
            Write-Verbose $Mailobj
        }
    }
}



$Names = get-mailbox 
Get-mailboxsizes -Name $Names -Verbose | export-csv "$env:userprofile\desktop\Powershell Reports\MailboxSizes.csv" -NoTypeInformation
