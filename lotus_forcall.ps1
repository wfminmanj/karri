<#
.Synopsis
Decline Category Review SharePoint List Items with PendingAction = 'Decline'
.DESCRIPTION
Set ReviewStatus = 'Declined'
Reset PendingAction = 'None'
Generate a regrets email for the submitter
.EXAMPLE
.\lotus_forcall.ps1 -Parmfile Lotus_Parms -Round 9
$psISE
#>

Param
 (
 # Name of local parameter file
 [Parameter(Mandatory=$true, Position=0)]
 [ValidateNotNull()] [ValidateNotNullOrEmpty()]
 $parmfile,

 # Required Arguement for Review Round
 [Parameter(Mandatory=$true, Position=1)]
 [ValidateNotNull()] [ValidateNotNullOrEmpty()]
 [ValidateRange(1,10)][Int]
 $round	)

 #### Setup ####
<#
Set-Location c:\psg
$parmfile = 'Lotus_Parms'
$round = 8
#>

$P = Get-Content ".\$parmfile.pson" | Out-String | Invoke-Expression

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

# $strURL   = 'http://sites/global/Grocery/CategoryReview/_vti_bin/Lists.asmx?WSDL'
# $guidReceived = '{9FCAECAE-11DF-498E-A539-D208AE7A348D}'
$strURL = $P.SPURI
$guidList = $P.SPListGUID

$re = "[{0}]" -f ([RegEx]::Escape([String][System.IO.Path]::GetInvalidFileNameChars()))

Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"
$outlook    = New-Object -ComObject Outlook.Application
$namespace  = $outlook.GetNameSpace("MAPI")

Function Get-SPCRForCall($url, $guid) {
  $rowLimit = "999"

  ## Begin CAML Select Statement
  ## Assemble a CAML Query that selects items "for call"
  $xmlDoc = New-Object System.Xml.XmlDocument
  $xmldecl = $xmlDoc.CreateXmlDeclaration("1.0", "utf-8", $null)
  $camlView = $xmlDoc.CreateElement("View")
  
  $elemntQuery    = $xmlDoc.CreateElement("Query")
  $elemntWhere    = $xmlDoc.CreateElement("Where")
  $elemntAnd      = $xmlDoc.CreateElement("And")
  $elemntAction   = $xmlDoc.CreateElement("Eq")
  $elemntRound    = $xmlDoc.CreateElement("Eq")
  
      #Action Criterion
          $elemntFRAction = $xmlDoc.CreateElement("FieldRef")
          $elemntFRAction.SetAttribute("Name","PendingAction")
      $elemntAction.AppendChild($elemntFRAction)   | Out-Null
          $elemntValueA   = $xmlDoc.CreateElement("Value")
          $elemntValueA.SetAttribute("Type","Choice")
          $elemntValueA.InnerText = 'ForCall'
      $elemntAction.AppendChild($elemntValueA)    | Out-Null
  
      #Round Criterion
          $elemntFRRound  = $xmlDoc.CreateElement("FieldRef")
          $elemntFRRound.SetAttribute("Name","ReviewRound")
      $elemntRound.AppendChild($elemntFRRound)   | Out-Null
          $elemntValueR   = $xmlDoc.CreateElement("Value")
          $elemntValueR.SetAttribute("Type","Text")
          $elemntValueR.InnerText = 'Round ' + $round
      $elemntRound.AppendChild($elemntValueR)   | Out-Null
  
  $elemntAnd.AppendChild($elemntAction)  | Out-Null
  $elemntAnd.AppendChild($elemntRound)   | Out-Null
  $elemntWhere.AppendChild($elemntAnd)   | Out-Null
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
  #$xmlDoc.Save("$([System.Environment]::GetEnvironmentVariable('TMP','MACHINE'))\se2_forcall.xml")
  ## End CAML Select Statement

  $objSelectProxy = New-WebServiceProxy -Uri $strURL  -Namespace SpWs  -UseDefaultCredential
  $ndReturnSE2 = $objSelectProxy.GetListItems($guid, $null, $elemntQuery, $elemntViewFld, $rowLimit, $elemntQueryOpt, $null)
  Return $ndReturnSE2
  }

Function Pause-Script {
    param([string] $pauseKey,
            [ConsoleModifiers] $modifier,
            [string] $prompt,
            [bool] $hideKeysStrokes)
             
    Write-Host -NoNewLine "Press $prompt to continue . . . "
    Do
    {
        $key = [Console]::ReadKey($hideKeysStrokes)
    } 
    While(($key.Key -ne $pauseKey) -or ($key.Modifiers -ne $modifer))   
     
    Write-Host
}

#### Get List Items with PendingAction = 'Decline' ####
$banana = Get-SPCRForCall $strURL $guidList
$orange = @()
If ($banana.Data.Row) {
If ($banana.Data.Row.Count) { Write-Output "Selected for task: $($banana.Data.Row.Count)"
       ForEach ( $m in $banana.Data.Row ) { $orange += $m } }
Else { Write-Output "Selected for task: 1" 
       $orange += $banana.Data.Row } }
Else { Write-Output "Selected for task: NONE"
       Break }

$PowerPointApp = New-Object -ComObject PowerPoint.Application
#$Presentation = $PowerPointApp.Presentations.Open($deck.FullName,$null,$null,[Microsoft.Office.Core.MsoTriState]::msoTrue)


$trmmodSHFT = [ConsoleModifiers]::Shift
$trmmodCTRL = [ConsoleModifiers]::Control
$trmmodALT = [ConsoleModifiers]::Alt
ForEach ( $m in $banana.Data.Row ) {
$msgkey = [System.Security.SecurityElement]::Escape(($m.ows_Subject -Replace $re, '-'))
$nodemsg = $X.SelectSingleNode("/repo/messages/msg[@msgkey='$($msgkey)']")
If ( $nodemsg ) {
        $alist = ""
        ForEach ( $j in $nodemsg.attachments.file ) { 
        If ( $j.Name -match '\.ppt?.' ) { 
        Write-Host -BackgroundColor Green -ForegroundColor Black "Trying to open $($j.Name)"
        $Presentation = $PowerPointApp.Presentations.Open($($repopath.FullName+"\"+$j.Name),$null,$null,[Microsoft.Office.Core.MsoTriState]::msoTrue) } }
        # Pause-Script "G" $trmmodCTRL "Ctrl + G" $true
        $Shell = New-Object -ComObject "WScript.Shell"
        $Button = $Shell.Popup("Click OK to continue.", 0, "Script Paused", 0)
               }
}



#### Wrapup ####


####KTHXBYE####
