# Description: Interactive PST export of Exchange 2012 mailboxes -wdb 15jun12
# status: prompts tested, rights assignment tested, mailbox export untested
ECHO "INS Exchange 2010 Exmerge "
echo "Did you give yourself the correct roles?"
echo "EG ""New-ManagementRoleAssignment -Role ""Mailbox Import Export"" -User INS\wbeam"""
do
{
    #get the mailbox name
    $MailBox = read-host -prompt "Type Mailbox username"
    #test name, if no good, loop back to get new name, if good continue
    $MBtest = get-user -Identity $MailBox
}
while ($MBtest.RecipientType -ne "UserMailBox")

ECHO "ok thats a mailbox"

# get the path to export
$PSTPath = read-host -prompt "Type export destination(EG \\server\The Share Path\test.pst)"
# test write permissions to path, if no good, loop back to get new path, if good continue
# Command HERE
# perform the actual export

New-MailboxExportRequest -Mailbox $MailBox -FilePath $PSTPath
ECHO "monitor!"
echo "will update status every 10 seconds"
#loop until request finishes
do
{
    $progress = Get-MailboxExportRequest
    #wait 60seconds
    Start-Sleep -Seconds 10
    #print output
    echo $progress.Status
}
while ($progress.status -ne "Completed")
 
ECHO "Done. Stopping the exmerge request"
Get-MailboxExportRequest | Remove-MailboxExportRequest
ECHO $MailBox " export to "$PSTPath "is Finished!" 
#EOF
