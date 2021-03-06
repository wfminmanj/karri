<#
.Synopsis
Decline Category Review SharePoint List Items with PendingAction = 'Decline'
.DESCRIPTION
Set ReviewStatus = 'Declined'
Reset PendingAction = 'None'
Generate a regrets email for the submitter
.EXAMPLE
.\lotus_decline.ps1
#>

Param
 (
 # Name of local parameter file
 [Parameter(Mandatory=$true, Position=0)]
 [ValidateNotNull()] [ValidateNotNullOrEmpty()]
 $parmfile,

 # Required Arguement for Usage Scenario
 [Parameter(Mandatory=$true, Position=1)]
 [ValidateNotNull()] [ValidateNotNullOrEmpty()]
 [ValidateSet('MsgSend','MsgDisplay','SPOnly')]
 $usagemode	)

 #### Setup ####
<#
Set-Location c:\psg
$parmfile = 'Lotus_Parms_DEV'
$parmfile = 'Lotus_Parms'
$usagemode = 'MsgDisplay'
#>

$P = Get-Content ".\$parmfile.pson" | Out-String | Invoke-Expression

$Logfile = $([System.Environment]::GetEnvironmentVariable('TMP','MACHINE')) + "\$($P.logfile_name).md"
If (!(Test-Path $Logfile)) {
   Try { Set-Content -Path $Logfile -Value ($null) }
   Catch { Write-Error "Logfile not found/valid: $($Logfile)"; Break } }
$frtime = Get-Date
Add-Content $Logfile -Value "###Script Start $($frtime.ToString('yyyy-MMM-dd_HH:mm:ss').ToUpper())$nl"

## Repo XML Existence Check
If (Test-Path $P.message_repo) {
$X = New-Object System.Xml.XmlDocument
$X.Load($P.message_repo)
$nodeMP = $X.SelectSingleNode('/repo/messages') 
$chkpath = $X.SelectSingleNode('/repo/path')
}
Else { Write-Host -Back Black -Fore Red "ERROR attempting to set repo as $($P.message_repo)" }


$repopath = $(Get-ChildItem $P.message_repo).Directory

# Repo XML Consistency Check
If ( -not $chkpath.'#text' -eq $repopath ) {
Write-Host -Back Black -Fore Red @"
Parameter Path: $($P.message_repo)
XML Repo Metadata Path: $($chkpath.'#text')
The XML Repo Metadata path found at the parameter path does not match!
This XML Repo Metadata may not be consistent with the stored files! 
"@ }

#$p.SPURI
#$strURL   = 'http://sites/global/Grocery/CategoryReview/_vti_bin/Lists.asmx?WSDL'
# $guidReceived = '{9FCAECAE-11DF-498E-A539-D208AE7A348D}'

# 2016-AUG-19 Jeff I: adding square brackets as problematic characters to the system invalid list
$xfn = "[{0}]" -f ([RegEx]::Escape([String][System.IO.Path]::GetInvalidFileNameChars()))
$re = $xfn + '|[\u0132-\u4000]|[\[\]]'
# $re = "[{0}]" -f ([RegEx]::Escape([String][System.IO.Path]::GetInvalidFileNameChars()))

Add-Type -AssemblyName System.Web
Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"
$outlook    = New-Object -ComObject Outlook.Application
$namespace  = $outlook.GetNameSpace("MAPI")
<#
Function Get-SPCRDecline($url, $guid) {
  $rowLimit = "999"

  ## Begin CAML Select Statement
  ## Assemble a CAML Query that selects items pending decline
  $xmlDoc = New-Object System.Xml.XmlDocument
  $xmldecl = $xmlDoc.CreateXmlDeclaration("1.0", "utf-8", $null)
  $camlView = $xmlDoc.CreateElement("View")
  $elemntQuery    = $xmlDoc.CreateElement("Query")

  $elemntWhere    = $xmlDoc.CreateElement("Where")
  $elemntEq       = $xmlDoc.CreateElement("Eq")
  $elemntFRLock   = $xmlDoc.CreateElement("FieldRef")
  $elemntFRLock.SetAttribute("Name","PendingAction")
  $elemntEq.AppendChild($elemntFRLock)   | Out-Null
  $elemntWhere.AppendChild($elemntEq)    | Out-Null
  $elemntQuery.AppendChild($elemntWhere) | Out-Null
  $elemntValue    = $xmlDoc.CreateElement("Value")
  $elemntValue.SetAttribute("Type","Choice")
  $elemntValue.InnerText = 'Decline'
  $elemntEq.AppendChild($elemntValue)    | Out-Null
  $elemntEq.AppendChild($elemntFRLock)   | Out-Null
  $elemntWhere.AppendChild($elemntEq)    | Out-Null
  $elemntQuery.AppendChild($elemntWhere) | Out-Null

  $camlView.AppendChild($elemntQuery) | Out-Null
  $elemntViewFld  = $xmlDoc.CreateElement("ViewFields")

  $elemntFRID     = $xmlDoc.CreateElement("FieldRef")
  $elemntFRID.SetAttribute("Name","ID")
  $elemntSubj     = $xmlDoc.CreateElement("FieldRef")
  $elemntSubj.SetAttribute("Name","Subject")
  $elemntViewFld.AppendChild($elemntSubj) | Out-Null
  $elemntViewFld.AppendChild($elemntFRID) | Out-Null
  $camlView.AppendChild($elemntViewFld)   | Out-Null

  $elemntQueryOpt = $xmlDoc.CreateElement("QueryOptions")
  $camlView.AppendChild($elemntQueryOpt) | Out-Null
  $xmlDoc.AppendChild($camlView)
  $xmlDoc.InsertBefore($xmldecl, $camlView) | Out-Null
  #$xmlDoc.Save("$([System.Environment]::GetEnvironmentVariable('TMP','MACHINE'))\se2_locked.xml")
  ## End CAML Select Statement

  $objSelectProxy = New-WebServiceProxy -Uri $url  -Namespace SpWs  -UseDefaultCredential
  $ndReturnSE2 = $objSelectProxy.GetListItems($guid, $null, $elemntQuery, $elemntViewFld, $rowLimit, $elemntQueryOpt, $null)
  Return $ndReturnSE2
  }
#>

Function Get-SPCRDecline($url, $guid) {
  $rowLimit = "999"

  ## Begin CAML Select Statement
  ## Assemble a CAML Query that selects items pending decline
  $xmlDoc = New-Object System.Xml.XmlDocument
  $xmldecl = $xmlDoc.CreateXmlDeclaration("1.0", "utf-8", $null)
  $camlView = $xmlDoc.CreateElement("View")
  $elemntQuery    = $xmlDoc.CreateElement("Query")

  $elemntWhere    = $xmlDoc.CreateElement("Where")
  $elemntEq       = $xmlDoc.CreateElement("Eq")
  $elemntFRLock   = $xmlDoc.CreateElement("FieldRef")
  $elemntFRLock.SetAttribute("Name","PendingAction")
  $elemntEq.AppendChild($elemntFRLock)   | Out-Null
  $elemntWhere.AppendChild($elemntEq)    | Out-Null
  $elemntQuery.AppendChild($elemntWhere) | Out-Null
  $elemntValue    = $xmlDoc.CreateElement("Value")
  $elemntValue.SetAttribute("Type","Choice")
  $elemntValue.InnerText = 'Decline'
  $elemntEq.AppendChild($elemntValue)    | Out-Null
  $elemntEq.AppendChild($elemntFRLock)   | Out-Null
  $elemntWhere.AppendChild($elemntEq)    | Out-Null
  $elemntQuery.AppendChild($elemntWhere) | Out-Null

  $camlView.AppendChild($elemntQuery) | Out-Null
  $elemntViewFld  = $xmlDoc.CreateElement("ViewFields")

  $elemntFRID     = $xmlDoc.CreateElement("FieldRef")
  $elemntFRID.SetAttribute("Name","ID")
  $elemntSubj     = $xmlDoc.CreateElement("FieldRef")
  $elemntSubj.SetAttribute("Name","Subject")
  $elemntLNK     = $xmlDoc.CreateElement("FieldRef")
  $elemntLNK.SetAttribute("Name","MessageLink")
  $elemntViewFld.AppendChild($elemntSubj) | Out-Null
  $elemntViewFld.AppendChild($elemntFRID) | Out-Null
  $elemntViewFld.AppendChild($elemntLNK)  | Out-Null
  $camlView.AppendChild($elemntViewFld)   | Out-Null

  $elemntQueryOpt = $xmlDoc.CreateElement("QueryOptions")
  $camlView.AppendChild($elemntQueryOpt) | Out-Null
  $xmlDoc.AppendChild($camlView)
  $xmlDoc.InsertBefore($xmldecl, $camlView) | Out-Null
  #$xmlDoc.Save("$([System.Environment]::GetEnvironmentVariable('TMP','MACHINE'))\se2_locked.xml")
  ## End CAML Select Statement

  $objSelectProxy = New-WebServiceProxy -Uri $url  -Namespace SpWs  -UseDefaultCredential
  $ndReturnSE2 = $objSelectProxy.GetListItems($guid, $null, $elemntQuery, $elemntViewFld, $rowLimit, $elemntQueryOpt, $null)
  Return $ndReturnSE2
  }

#### Get List Items with PendingAction = 'Decline' ####
$banana = Get-SPCRDecline $P.SPURI $P.SPListGUID
$orange = @()
If ($banana.Data.Row) {
If ($banana.Data.Row.Count) { Write-Output "Selected for Decline: $($banana.Data.Row.Count)"
       ForEach ( $m in $banana.Data.Row ) { $orange += $m } }
Else { Write-Output "Selected for Decline: 1" 
       $orange += $banana.Data.Row } }
Else { Write-Output "Selected for Decline: NONE"
       Break }

<########################################################################## Begin Main ##
 __   ___  __                              
|__) |__  / _` | |\ |     |\/|  /\  | |\ | 
|__) |___ \__> | | \|     |  | /~~\ | | \| 

########################################################################################>

# If ( $banana.Data.Row.Count -gt 0 ) {
  ## Begin CAML Update Batch Statement  $orange[0].ows_Subject
  ## Assemble a list of item ids into a batch
$xmlDoc = New-Object System.Xml.XmlDocument
$xmldecl = $xmlDoc.CreateXmlDeclaration("1.0", "utf-8", $null)
$camlBatch = $xmlDoc.CreateElement("Batch")

$n=1
$itemseq = 1
$msgerrors = 0
ForEach ( $i in $orange ) {

 #### Begin Decline Message #### $P
  Write-Progress -Activity "Processing List Items" -Status "$($i.ows_Subject) [ $n of $($orange.count) ]" -PercentComplete (100*($n/$orange.count))
  Write-Host -BackgroundColor Green -ForegroundColor Black "[[ $($i.ows_Subject) ]]" -NoNewline
  # $msgkey = [System.Security.SecurityElement]::Escape(($i.ows_Subject -Replace $re, '-'))
  # $i.ows_MessageLink
  $i.ows_MessageLink -match '(?<=\?msgkey=).*(?=,)' | Out-Null
  $msgkey = $Matches[0]
  $msghtmlxcd = [System.Web.HttpUtility]::HtmlDecode($msgkey)
  $nodemsg = $X.SelectSingleNode("/repo/messages/msg[@msgkey='$($msgkey)']")
        If ( $nodemsg ) {
              $alist = ""
              ForEach ( $m in $nodemsg.attachments.file ) { $alist += "<li>$($m.name)</li>" }
           }
           Else { Write-Host "Repo node not found for [$($i.ows_Subject)]" }
  $msg = Get-ChildItem $([System.Web.HttpUtility]::HtmlDecode($nodemsg.href)).SubString(8)
  Try { $peach = $outlook.CreateItemFromTemplate($msg) }
  Catch { $msgerrors ++ }
  If ( $peach ) {
                  $rail = $outlook.CreateItem(0)
                  $rail.SentOnBehalfOfName = $P.mailbox
                  $rail.To = $peach.SenderEmailAddress
                  $rail.Subject = 'RE: ' + $peach.Subject
                  $rail.HTMLBody = $P.decline_open + $alist + $P.decline_close + $rail.HTMLBody
                  Try {
                  If ( $usagemode -eq 'MsgDisplay' ) { $rail.Display() }
                  ElseIf ( $usagemode -eq 'MsgSend' ) { $rail.Send() } }
                  Catch { $msgerrors ++ } 
                  Write-Host ([char]60) -NoNewline
                  Add-Content $Logfile -Value "####$usagemode :: $msgkey $nl"
               }
               Else { Write-Host "Repo message not found for [$($i.ows_Subject)]" }

 ####  End  Decline Message ####

 #### Begin CAML Batch Build ####
  
  If ( $msgerrors -eq 0 ) {
      $camlItem = $xmlDoc.CreateElement("Method")
      $camlItem.SetAttribute("ID",$itemseq)
      $camlItem.SetAttribute("Cmd","Update")
      $camlID = $xmlDoc.CreateElement("Field")
      $camlID.SetAttribute("Name","ID")
      $camlID.InnerText = $i.ows_ID
      $camlItem.AppendChild($camlID) | Out-Null

      $camlAction = $xmlDoc.CreateElement("Field")
      $camlAction.SetAttribute("Name","PendingAction")
      $camlAction.InnerText = 'None'

      $camlStatus = $xmlDoc.CreateElement("Field")
      $camlStatus.SetAttribute("Name","ReviewStatus")
      $camlStatus.InnerText = 'Declined'
      $camlItem.AppendChild($camlStatus)  | Out-Null
      $camlItem.AppendChild($camlAction)  | Out-Null
      $camlBatch.AppendChild($camlItem) | Out-Null
      Write-Host ([char]62) -NoNewline
    }

####  End  CAML Batch Build ####

  $itemseq ++
  $n++
  Write-Host "    $itemseq"
  $msgerrors = 0
  $msg = $null
  $peach = $null
  }
  
If ($itemseq -gt 1 ) {
    Add-Content $Logfile -Value "$nl####CAML Batch $nl"
    $camlBatch.Method.Field | Where-Object { $_.Name -contains 'ID' } | Select-Object '#text' | Add-Content $Logfile
    $xmlDoc.AppendChild($camlBatch) | Out-Null
    $xmlDoc.InsertBefore($xmldecl, $camlBatch) | Out-Null

    $objDMLProxy = New-WebServiceProxy -Uri $P.SPURI  -Namespace SpWs  -UseDefaultCredential
    $ndReturnUP6 = $objDMLProxy.UpdateListItems($P.SPListGUID, $camlBatch) 
  }
Else { Write-Warning "The Batch is empty: No List Items have been updated." }

# $camlBatch.ChildNodes.Count


<##########################################################################  End  Main ##
    ___       __                          
   |__  |\ | |  \        |\/|  /\  | |\ | 
   |___ | \| |__/        |  | /~~\ | | \| 

########################################################################################>

<# A01 logic
If ( $usagemode -ne 'SPOnly') {
      ForEach ( $i in $orange ) {

            # Open the message from the repo and reply with the decline message
            $msgkey = [System.Security.SecurityElement]::Escape(($i.ows_Subject -Replace $re, '-'))
            $nodemsg = $X.SelectSingleNode("/repo/messages/msg[@msgkey='$($msgkey)']")
            If ( $nodemsg ) {
                  $alist = ""
                  ForEach ( $m in $nodemsg.attachments.file ) { $alist += "<li>$($m.name)</li>" }
               }
               Else { Write-Host "Repo node not found for [$($i.ows_Subject)]" }
            # Get-ChildItem "$repopath\$msgkey.msg"
              <# 
                 $msgkey = [System.Security.SecurityElement]::Escape(($banana.Data.Row.ows_Subject -Replace $re, '-'))
                 Get-ChildItem "$repopath\$msgkey.msg"
                 $msgkey = [System.Security.SecurityElement]::Escape(($i.Subject -Replace $re, '-'))
                #
            $peach = $outlook.CreateItemFromTemplate($(Get-ChildItem "$repopath\$msgkey.msg"))
            if ( $peach ) {
                  $rail = $outlook.CreateItem(0)
                  $rail.SentOnBehalfOfName = $P.mailbox
                  $rail.To = $peach.SenderEmailAddress
                  $rail.Subject = 'RE: ' + $peach.Subject
                  $rail.HTMLBody = $P.decline_open + $alist + $P.decline_close + $rail.HTMLBody
                  If ( $usagemode -eq 'MsgDisplay' ) { $rail.Display() }
                  ElseIf ( $usagemode -eq 'MsgSend' ) { $rail.Send() }
               }
               Else { Write-Host "Repo message not found for [$($i.ows_Subject)]" }
             
         }
   }
#>

Add-Content $Logfile -Value "$nl###Message Processing Complete  $($frtime.ToString('yyyy-MMM-dd_HH:mm:ss').ToUpper())$nl"

#### Wrapup ####

$rail = $null
$outlook = $null

####KTHXBYE####
