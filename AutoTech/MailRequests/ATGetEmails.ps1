﻿# Name of the mailbox to pull attachments from
$MailboxName = 'AutoTech@contoso.com'
 
# Location to move attachments
$downloadDirectory = '\\mailstore.contoso.com\c$\AutoTech\MailRequests'
 
# Path to the Web Services dll
$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"
[VOID][Reflection.Assembly]::LoadFile($dllpath)
 
# Create the new web services object
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1)
 
# Create the LDAP security string in order to log into the mailbox
$windowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$sidbind = "LDAP://<SID=" + $windowsIdentity.user.Value.ToString() + ">"
$aceuser = [ADSI]$sidbind
 
# Auto discover the URL used to pull the attachments
$service.AutodiscoverUrl($aceuser.mail.ToString())
 
# Get the folder id of the Inbox
$folderid = new-object  Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$MailboxName)
$InboxFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
 
# Find mail in the Inbox with attachments
$Sfha = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::HasAttachments, $true)
$sfCollection = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::And);
$sfCollection.add($Sfha)
 
# Grab all the mail that meets the prerequisites
$view = new-object Microsoft.Exchange.WebServices.Data.ItemView(2000)
$frFolderResult = $InboxFolder.FindItems($sfCollection,$view)
 
# Loop through the emails
foreach ($miMailItems in $frFolderResult.Items){
 
    # Load the message
    $miMailItems.Load()
 
    # Loop through the attachments
    foreach($attach in $miMailItems.Attachments){
 
        # Load the attachment
        $attach.Load()
 
        # Save the attachment to the predefined location
        $fiFile = new-object System.IO.FileStream(($downloadDirectory + “\” + (Get-Date).Millisecond + "_" + $attach.Name.ToString()), [System.IO.FileMode]::Create)
        $fiFile.Write($attach.Content, 0, $attach.Content.Length)
        $fiFile.Close()
    }
 
    # Mark the email as read
    $miMailItems.isread = $true
    $miMailItems.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite)
 
# Move email to Processed folder
$FolderName = "Processed"
$oFolder = new-object Microsoft.Exchange.WebServices.Data.Folder($Service)
$oFolder.DisplayName = $FolderName

#Define Folder View really only want to return one object
$fvFolderView = new-object Microsoft.Exchange.WebServices.Data.FolderView(1)

#Define a Search folder that is going to do a search based on the DisplayName of the folder
$SfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,$FolderName)

#Get Inbox ID
$InboxID = $Inboxfolder.ID

#Do the Search
$findFolderResults = $Service.FindFolders($InboxID,$SfSearchFilter,$fvFolderView)

If ($findFolderResults.TotalCount -ne 0)
{	
	$DestID = $findFolderResults.Folders[0].Id.UniqueID
	$miMailItems.Move("DeletedItems") | Out-Null
}
 
 
 
    # Delete the message (optional)
    #[VOID]$miMailItems.Move("DeletedItems")
}
