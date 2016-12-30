<#
.DESCRIPTION
   Take a template OutLook message from the Drafts folder
   Generate an email messsge for each line in a CSV file having columns:
          - dlist as a semicolon delimited list of email addresses
          - atchmnt as a fullpath to a workbook to attach
          - brand as a brand ID - not used
.EXAMPLE
   .\MSOutlookMessage_StaticBody.ps1 MSOutlook_Parms_KB MsgSend
.NOTES
    The parm file [subject] parameter will be matched to the message [ConversationTopic], 
        which does not include colon delimited prefixes like 'RE:' or 'FW:'
    Example parm file contents
         @{
              'mailbox' = 'Kate.Brunson@wholefoods.com'
              'subject' = 'Avalons'
              'csvpath' = '\\wfm-team\Team\RegionalPurchasing\National Promotions\WFMPromotions\CSVFiles\foo.csv'
            }


#>

<####### Setup ###########
  ___ ___ _____ _   _ ___
 / __| __|_   _| | | | _ \
 \__ \ _|  | | | |_| |  _/
 |___/___| |_|  \___/|_|
##########################>

#### Set Parameters ####

Param
 (
 # Name of local parameter file
 [Parameter(Mandatory=$true, Position=0)]
 [ValidateNotNull()] [ValidateNotNullOrEmpty()]
 $parmfile,

 # Required Arguement for Usage Scenario
 [Parameter(Mandatory=$true, Position=1)]
 [ValidateNotNull()] [ValidateNotNullOrEmpty()]
 [ValidateSet('MsgSend','MsgDisplay','Debug')]
 $usagemode		)

# $usagemode = 'MsgDisplay'
# $parmfile  = 'MSOutlook_Parms_JI'
If ( -not $(Test-Path "c:\psg\$parmfile.pson") ) {
      Write-Host -BackgroundColor Magenta -ForegroundColor Black `
      "!! Parameter File not found: c:\psg\$parmfile.pson !!"
      Break }
Try {
      $P = Get-Content "c:\psg\$parmfile.pson" | Out-String | Invoke-Expression }
Catch {
      Write-Host -BackgroundColor Magenta -ForegroundColor Black `
      "!! Parameter File Failed: c:\psg\$parmfile.pson !!"
      Break }
# $P | Format-Table -AutoSize

#### Open/Connect-To MS Outlook ####
Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"
$olApp     = New-Object -ComObject Outlook.Application
$namespace = $olApp.GetNameSpace("MAPI")

#### List Folder Names ####
# $namespace.Folders | Select-Object -Property Name

#### Set the working folder names ####
# $fn_parent = $P.mailbox
$fn_source = 'Drafts'

#### Set the woking folder objects ####
$parent = $($namespace.Folders | Where-Object {$_.Name -eq $P.mailbox})
$source = $($parent.Folders | Where-Object {$_.Name -eq $fn_source})

# $tmplt_subject = $P.subject
# $source.Items | Where-Object {$_.ConversationTopic -match $P.subject} | Measure-Object
# $msg_tmplt = $source.Items | Where-Object {$P.subject -match $_.ConversationTopic}
$msg_tmplt = $source.Items | Where-Object {$_.Categories -match $P.MessageTag}

If ( -not $msg_tmplt) {
        Write-Host -BackgroundColor Magenta -ForegroundColor Black "!! No subject match on $($P.subject) !!"
        Break }
ElseIf ( $msg_tmplt.Count -gt 1 ) {
        Write-Host -BackgroundColor Magenta -ForegroundColor Black "!! Multiple matches on $($P.subject) !!"
        Break }

# message object check:    $msg_tmplt.Display()

#### Load Message Parameters ####

If ( $P.csvhdrs ) { $list = $(Import-CSV $P.csvpath) }
Else { $list = $(Import-CSV $P.csvpath -Header dlist,FullName,BrandID) }
Write-Host -BackgroundColor DarkBlue -ForegroundColor White "+++ List Count: $($list.Count)"
If ( $list.Count -gt 1 ) {
        Write-Host -BackgroundColor DarkBlue -ForegroundColor White "--- First Item: "
        $list[0]  | Format-List *
        Write-Host -BackgroundColor DarkBlue -ForegroundColor White "--- Last Item: "
        $list[-1] | Format-List * }
$summary = New-Object -TypeName PSObject
$summary | Add-Member -M NoteProperty -Name rows_in -Val $list.Count
$summary | Add-Member -M NoteProperty -Name err_send -Val 0

<######## Main #########
  __  __   _   ___ _  _
 |  \/  | /_\ |_ _| \| |
 | |\/| |/ _ \ | || .` |
 |_|  |_/_/ \_\___|_|\_|
########################>

#### Begin Generate Messages ####

$timer = [Diagnostics.Stopwatch]::StartNew()
ForEach ($i in $list ) {
        $banana = $msg_tmplt.Copy()
        $banana.SentOnBehalfOfName = $P.From
        $i.dlist.TrimEnd([char]59).Split([char]59) | ForEach-Object { $banana.Recipients.Add($_) | Out-Null }
    Try {
        $banana.Attachments.Add($i.FullName) | Out-Null }
    Catch {
        Write-Host -BackgroundColor Magenta -ForegroundColor Black "!! Attachment Failed: $($i.FullName.Split([char]92)[-1]) !!" }
    Try {
        If ( 'MsgDisplay' -Contains $usagemode ) { $banana.Display() }
        ElseIf ( 'MsgSend' -Contains $usagemode ) { $banana.Send() }
        Write-Host -BackgroundColor Green -ForegroundColor Black "[[ $($i.FullName.Split([char]92)[-1]) ]]" }
    Catch {
        $banana.Save()
        $summary.err_send += 1
        Write-Host -BackgroundColor Magenta -ForegroundColor Black "!! $($i.bid) $($i.dlist) !!" }
  }
$timer.Stop()

#### End Generate Messages ####

<############# Wrapup ##############
 __      _____    _   ___ _   _ ___
 \ \    / / _ \  /_\ | _ \ | | | _ \
  \ \/\/ /|   / / _ \|  _/ |_| |  _/
   \_/\_/ |_|_\/_/ \_\_|  \___/|_|
####################################>

#### Write stats to the console ####
$summary | Add-Member -M NoteProperty -Name minutes -Val $($timer.Elapsed | Select-Object -Property Minutes).Minutes
$summary | Add-Member -M NoteProperty -Name seconds -Val $($timer.Elapsed | Select-Object -Property Seconds).Seconds
$summary | Format-Table -AutoSize

Remove-Variable olApp

####KTHXBYE####
