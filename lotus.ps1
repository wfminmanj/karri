<#
.Synopsis
Outlook Email Message Ingestion
.DESCRIPTION
Outlook Email Message Ingestion - Global Grocery Category Review Process 2015
.VERSION
A03
.EXAMPLE
.\lotus.ps1 Lotus_Parms_JI FileOnly
Where the parm file is in pson format; example contents:
@{
'SPURI' = 'http://sites/global/Grocery/CategoryReview/_vti_bin/Lists.asmx?WSDL'
'SPListGUID' = '{6ABF96D9-197E-4ED9-AC10-B17C3C9CEC23}'
'SPViewGUID' = '{154A1E10-E8E8-4DD3-91DC-9B283663DA3D}'
'RowLimit' = "999"
'SourceFolder' = 'inbound_response'
'TargetFolder' = 'inbound_archive'
'OLEDBProvider' = 'Provider=Microsoft.ACE.OLEDB.12.0'
'OLEDBExtend' = 'Extended Properties=Excel 12.0'
'ExportTag' = 'ps1_Export'
'logfile_name' = 'lotusdev'
'linksrv' = 'http://localhost:8090/linklkup?msgkey='
'message_repo' = '\\wfm-team\team\RegionalPurchasing\National Promotions\_dev_repo_\_dev_repo_.xml'
'mailbox' = 'GlobalGrocery.Promotions@wholefoods.com'
'prepend' = 'FYI: '
'opening' = @"
<p  style="font-family:sans-serif">Greetings,<br>The following files were parsed from this message:</p>
<ul style="font-family:monospace">
"@
'closing' = @"
</ul><br/><p style="font-family:sans-serif">&gt_ by a Procurement Non-Perishables Analytics PowerShell script</p>
<p style="font-family:sans-serif"><span style="color:blue">Whole Foods Market | Global Headquarters<br/>
550 Bowie Street | </span><span style="color:darkred">Austin, TX 78703</span><br/>
<span style="color:blue">P# 512.477.5566 | F# 512.499.6593</span></p>
"@
}
.INPUTS
1 of 2 Name of local parameter file
2 of 2 Usage Mode [ MsgSend sends reply messages, MsgDisplay opens reply messages, FileOnly spools messaging data to output
.OUTPUTS
Markdown formatted logfile
.NOTES
Requires MS Outlook (email messaging engine for mailbox message management and message generation)
Requires SharePoint (user list/library)
Requires a file system directory as repository where email files will be stored and an xml record will be maintained
Requires a log file in the system TEMP directory
to parse markdown syntax on Windows OS, try MarkdownPad 2, further options at http://mashable.com/2013/06/24/markdown-tools/
#### TODO ####
+ Parameterize
+ Error handling on ZIP extraction with unicode filenames
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
 [ValidateSet('MsgSend','MsgDisplay','FileOnly')]
 $usagemode		)

<#  Functions
 ___            __  ___    __        __
|__  |  | |\ | /  `  |  | /  \ |\ | /__`
|    \__/ | \| \__,  |  | \__/ | \| .__/
#>
Function Recurse-Folders ( $obj, $str ) {
      ForEach ( $f in $obj.Folders ) {
            If ( $str -Contains $f.Name ) { Return $f }
            ElseIf ( $f.Folders.Count -gt 0 -and $f.Name -NotLike 'Public*'  ) { Recurse-Folders $f $str }
    }      }
Function Get-Checksum ($file, $crypto_provider) {
	If ($crypto_provider -eq $null) {
		$crypto_provider = New-Object -TypeName System.Security.Cryptography.MD5CryptoServiceProvider
	  }
	$file_info	= Get-Item $file
	Trap {
	    Continue } $stream = $file_info.OpenRead()
	If ($? -eq $false) {
		Return $null
	  }
	$bytes		= $crypto_provider.ComputeHash($stream)
	$checksum	= ''
	ForEach ($byte in $bytes) {
		$checksum	+= $byte.ToString('x2').ToUpper()
	  }
	$stream.Close() | Out-Null
	Return $checksum
 }
Function Get-SomeCheck {
 Param ([string]$strFoo)

 $objMD5 = New-Object -TypeName System.Security.Cryptography.MD5CryptoServiceProvider
 $objENC = New-Object -TypeName System.Text.UTF8Encoding
 $bash = $objMD5.ComputeHash($objENC.GetBytes($strFoo))

 $checksum = ''
 ForEach ($byte in $bash) { $checksum	+= $byte.ToString('x2').ToUpper() }
 $checksum
 }
Function Get-UnqFilePath ( $pstr, $nstr) {
 If ( Test-Path "$($pstr)\$($nstr)" ) {
 $arr = $nstr.Split([Char]46)
 $s1 = $arr[0..($arr.length - 2)]
 $cx = $( Get-ChildItem "$($pstr)\$($s1 -Join [Char]46)*" ).Name
 $s1 += Get-SomeCheck $cx
 $s1 += $arr[-1]
 "$($pstr)\$($s1 -Join [Char]46)" }
 Else  { "$($pstr)\$($nstr)" }  }

<# Setup
 __   ___ ___       __
/__` |__   |  |  | |__)
.__/ |___  |  \__/ |
#>
# $usagemode ValidateSet('Debug','Test','Normal')

Add-Type -AssemblyName "System.IO.Compression.FileSystem"
Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"
$binding = "System.Reflection.BindingFlags" -as [Type]

$P = Get-Content ".\$parmfile.pson" | Out-String | Invoke-Expression

If (Test-Path $P.message_repo) {
$X = New-Object System.Xml.XmlDocument
$X.Load($P.message_repo)
$nodeMP = $X.SelectSingleNode('/repo/messages') 
$chkpath = $X.SelectSingleNode('/repo/path')
}
Else { Write-Host -Back Black -Fore Red "ERROR attempting to set repo as $($P.message_repo)" }

$repopath = $(Get-ChildItem $P.message_repo).Directory

If ( -not $chkpath.'#text' -eq $repopath ) {
Write-Host -Back Black -Fore Red @"
Parameter Path: $($P.message_repo)
XML Repo Metadata Path: $($chkpath.'#text')
The XML Repo Metadata path found at the parameter path does not match!
This XML Repo Metadata may not be consistent with the stored files! 
"@ }

# $usagemode = 'FileOnly'
 ## $usagemode = 'FileOnly'
 ## $parmfile  = 'Hadar_Parms_JI'
 # $P = Get-Content ".\$parmfile.pson" | Out-String | Invoke-Expression

<# Note on retrieval of HTML param tag seeded values
$body = "<param PromoBrandName='Mr Slurry';MerchCategory='RTD Coffee' />"
If ($body -Match "(?<=(<param.*PromoBrandName\=['""]))(?<iBN>[^'""]*)(?=['""])") {$Matches['iBN']}
If ($body -Match "(?<=(<param.*MerchCategory\=['""]))(?<iMC>[^'""]*)(?=['""])") {$Matches['iMC']}
#>

$FormVersions = @( 'V299909', 'V201601' )
$NIQuery = @"
Select [LaunchType], [hCategory], [Brand], count(*) as [n]
From [New Items&ACV Gaps$]
Where [Eval] = TRUE and [UPC] <> '0900000000000'
Group By [LaunchType], [hCategory], [Brand]
"@

$Logfile = $([System.Environment]::GetEnvironmentVariable('TMP','MACHINE')) + "\$($P.logfile_name).md"
If (!(Test-Path $Logfile)) {
   Try { Set-Content -Path $Logfile -Value ($null) }
   Catch { Write-Error "Logfile not found/valid: $($Logfile)"; Break } }
$nl = [char]13
$sc = '~sample~'
$se = '~extrct~'
$re = "[{0}]" -f ([RegEx]::Escape([String][System.IO.Path]::GetInvalidFileNameChars()))
$frtime = Get-Date
$moved = @()
$attachmentcount = 0
$msgfilecount = 0
If ( -Not (Test-Path("$repopath\$sc"))  ) { New-Item -ItemType Directory -Force -Path "$repopath\$sc" }
If ( -Not (Test-Path("$repopath\$se"))  ) { New-Item -ItemType Directory -Force -Path "$repopath\$se" }
Add-Content $Logfile -Value "###Script Start $($frtime.ToString('yyyy-MMM-dd_HH:mm:ss').ToUpper())$nl"
Add-Content $Logfile -Value "#####Usage Mode: $($usagemode)$nl"
@"
 __        __              ___ ___  ___  __   __
|__)  /\  |__)  /\   |\/| |__   |  |__  |__) /__``
|    /~~\ |  \ /~~\  |  | |___  |  |___ |  \ .__/
"@
$P | Format-Table -AutoSize

# Get SharePoint Object as $objSP
Try{ $objSP = New-WebServiceProxy -Uri $P.SPURI  -Namespace SpWs  -UseDefaultCredential }
Catch{ Write-Error $_ -ErrorAction:'SilentlyContinue' }





<# Read Source
 __   ___       __      __   __        __   __   ___
|__) |__   /\  |  \    /__` /  \ |  | |__) /  ` |__
|  \ |___ /~~\ |__/    .__/ \__/ \__/ |  \ \__, |___
#>

# Get Outlook Application objects
$olFolders = "Microsoft.Office.Interop.Outlook.olDefaultFolders" -as [type]
$olSaveType = "Microsoft.Office.Interop.Outlook.OlSaveAsType" -as [type]
$olClass = "Microsoft.Office.Interop.Outlook.OlObjectClass" -as [type]
$olApp     = New-Object -ComObject Outlook.Application
$namespace = $olApp.GetNameSpace("MAPI")

# Get objects for the Outlook folders
Try { $oltrgfolder = $( Recurse-Folders $namespace $P.TargetFolder )[0] }
Catch { Write-Host -Back Black -Fore Red "ERROR attempting to match source folder name $($P.TargetFolder)" }
Try { $olsrcfolder = $( Recurse-Folders $namespace $P.SourceFolder )[0] }
Catch { Write-Host -Back Black -Fore Red "ERROR attempting to match source folder name $($P.SourceFolder)" }

## When selecting the Inbox by setting the SourceFolder parameter to a mailbox/email
If ($olsrcfolder.Folders | Where-Object { 'Inbox' -Contains $_.Name }) {
$olsrcfolder = $olsrcfolder.Folders | Where-Object { 'Inbox' -Contains $_.Name }}

# Get a collection of messages from the source folder
$colmsg = $olsrcfolder.Items | Where-Object {$_.Categories -match $P.ExportTag}

If ($colmsg.Count -gt 1 ) { $cardcount = $colmsg.Count }
Else { $cardcount = 1 }

<#
$colmsg | Select-Object -Property ConversationTopic,ConversationID | Format-Table -AutoSize
#>

<# Begin Message Loop
 __   ___  __                   ___  __   __        __   ___          __   __   __
|__) |__  / _` | |\ |     |\/| |__  /__` /__`  /\  / _` |__     |    /  \ /  \ |__)
|__) |___ \__> | | \|     |  | |___ .__/ .__/ /~~\ \__> |___    |___ \__/ \__/ |
#>
$card = 1
ForEach ( $i in $colmsg ) {
Write-Progress -Activity "Processing Messages" -Status "Loading $card of $cardcount" -PercentComplete (100*$card/$cardcount)
      

$MG = New-Object –TypeName PSObject
$MG | Add-Member –M NoteProperty –Name attachments –Val $i.Attachments.Count
If ( $MG.attachments -gt 0) {
$MG | Add-Member –M NoteProperty –Name zipattachments –Val 0
$MG | Add-Member –M NoteProperty –Name sigattachments –Val 0
$MG | Add-Member –M NoteProperty –Name savedfiles –Val 0
$MG | Add-Member –M NoteProperty –Name skippedfiles –Val 0
}
If ( -not ($usagemode -eq  'FileOnly') ) { $abody = $P.opening }
$nodemsg = $null
$xlfiles = @()
$brandguess = 'TBD'
$categoryguess = 'TBD'

# Get the Sender Email as $sender
If ( $i.SenderEmailType -eq 'EX' ) {
     $exuser = $olApp.Session.GetGlobalAddressList().AddressEntries.Item($i.Sender.Name)
     $sender = $exuser.GetExchangeUser().PrimarySmtpAddress }
   Else { $sender = $i.SenderEmailAddress }

# Create a key value for the message based on Conversation Topic
## $msgkey = [System.Security.SecurityElement]::Escape(($i.ConversationTopic -Replace $re, '-'))
If ($i.ConversationTopic.Length -gt 5 ) {
$msgkey = $i.ConversationTopic -Replace '[\W-[\.]]', '-' }
Else {$msgkey = "Empty-Subject-($($i.ConversationTopic -Replace '[\W-[\.]]', '-'))-From-$($sender -Replace '[\W-[\.]]', '-')"}
#-#}

# Create a file name based on the subject; replace spaces so it makes a readable URL
## $msgfn = $i.Subject -Replace $re, '-'
$msgfn = $msgkey
$i.SaveAs("$repopath\$sc\$msgfn.msg", $olSaveType::olMSG)
$fsimage = $olApp.CreateItemFromTemplate("$repopath\$sc\$msgfn.msg") 

# Get a review round from the message categories
$arycat = $($i.Categories -Split [char]44 | ForEach{ $_.Trim()})
$outcat = $arycat | Sort-Object | Where-Object { $_ -notmatch $P.ExportTag }
$rnd = $arycat | Sort-Object | Where-Object { $_ -match 'round' }
# Update categories less the export tag
$i.Categories = $outcat -Join [char]44
If ($rnd) { If ( $rnd.Count -gt 1 ) { $rnd = $rnd[0] } }
Else { $rnd = 'Round_10' }

# Apply the message to the file system
If ( Test-Path "$repopath\$msgfn.msg" ) {
      $mcount = 0
      #When the file name already exists, check if the contents are also duplicated
      ForEach ( $n in Get-ChildItem "$repopath\$msgfn*" ) {
            $compareimage = $olApp.CreateItemFromTemplate($n.FullName)
            If ( $fsimage.SentOn -eq $compareimage.SentOn -and $fsimage.Body -eq $compareimage.Body ) {
            $mcount ++ }  }
      If ( $mcount -gt 0 ) {
            #When there are duplicated contents, don't save an additional copy
            $MG | Add-Member –M NoteProperty –Name fs –Val   'skip-duplicate'
            Remove-Item "$repopath\$sc\$msgfn.msg" }
         Else {
              #When the contents are new, save the file with a new name
              $MG | Add-Member –M NoteProperty –Name fs –Val 'save-revision'
              $nom = Get-UnqFilePath $repopath "$msgfn.msg"
              Move-Item "$repopath\$sc\$msgfn.msg" -Destination $nom
              Set-ItemProperty $nom -Name IsReadOnly -Value $true
              $msglkup = Get-Item $nom
              Write-Host -ForeGroundColor Magenta "----> $($msglkup.Name) [$($msglkup.Length)_B]"
              Add-Content $Logfile -Value "  - **$($i.Subject)**, $($msglkup.Length)_Bytes, at $($msglkup.FullName)"
              $MG.savedfiles ++
           } }
   Else { # Save Message file
           $MG | Add-Member –M NoteProperty –Name fs –Val    'save-new'
           Move-Item "$repopath\$sc\$msgfn.msg" -Destination "$repopath\$msgfn.msg"
           Set-ItemProperty "$repopath\$msgfn.msg" -Name IsReadOnly -Value $true
           $msglkup = Get-Item "$repopath\$msgfn.msg"
           Write-Host -ForeGroundColor Magenta "----> $($msglkup.Name) [$($msglkup.Length)_B]"
           Add-Content $Logfile -Value "  - **$($i.Subject)**, $($msglkup.Length)_Bytes, at $($msglkup.FullName)"
   }

# Apply the message to the repo XML
$nodemsg = $X.SelectSingleNode("/repo/messages/msg[@msgkey='$($msgkey)']")
If ( $nodemsg ) { # When a node for the message key already exists, compare the .msg file checksum to the child/file nodes
      If ( Test-Path "$repopath\$msgfn.msg" ) {
      $orange = Get-Item "$repopath\$msgfn.msg"
      $mchksum = Get-Checksum $orange.FullName
      $nodefile = $X.SelectSingleNode("/repo/messages/msg[@msgkey='$($msgkey)']/attachments/file[@chksum='$($mchksum)']")
      If ( $nodefile ) { # When a file node with the checksum already exists, don't append a node
         $MG | Add-Member –M NoteProperty –Name xml –Val 'skip-duplicate' }
      Else { # Append a file node
         $MG | Add-Member –M NoteProperty –Name xml –Val   'append-filenode'
         $nodeattach = $X.SelectSingleNode("/repo/messages/msg[@msgkey='$($msgkey)']/attachments")
         If ( -not $nodeattach ) {
            $nodeattach = $X.CreateElement('attachments')
            $nodemsg.AppendChild($nodeattach) | Out-Null }
         $nodefile = $X.CreateElement('file')
         $nodefile.SetAttribute('name',$orange.Name)
         $nodefile.SetAttribute('lmdt',$orange.LastWriteTime.ToString('s'))
         $nodefile.SetAttribute('bytes',$orange.Length)
         $nodefile.SetAttribute('chksum',$mchksum)
         $nodefile.SetAttribute('href',$([System.Security.SecurityElement]::Escape('file:///' + "$($msglkup.FullName)")))
         $nodeattach.AppendChild($nodefile) | Out-Null } }
    Else { Write-Warning "Message Node Points to NULL file!" }
 }
Else {    # Append a message node
   $MG | Add-Member –M NoteProperty –Name xml –Val 'append-messagenode'
   $nodemsg = $X.CreateElement('msg')
   $nodemsg.SetAttribute('subject',$([System.Security.SecurityElement]::Escape($i.Subject)))
   $nodemsg.SetAttribute('sender',$sender)
   $nodemsg.SetAttribute('sentdate',$i.SentOn.ToString('s'))
   $nodemsg.SetAttribute('href',$([System.Security.SecurityElement]::Escape('file:///' + "$repopath\$msgfn.msg")))
   $nodemsg.SetAttribute('msgkey',$msgkey)
   $nodeMP.AppendChild($nodemsg) | Out-Null
} # End Append a message node

If ( $MG.attachments -gt 0 ) {
   # Append nodes for new attachments
   $nodeattach = $X.SelectSingleNode("/repo/messages//msg[@msgkey='$($msgkey)']/attachments")
   If ( -not $nodeattach ) {
   $nodeattach = $X.CreateElement('attachments')
   $nodemsg.AppendChild($nodeattach) | Out-Null }
   ForEach ( $a in $i.Attachments ) {
   If ( $a.FileName -match '.zip?$' ) { # Begin Zip Files
    $MG.zipattachments ++
    $zipName = ($a.FileName -Replace '[\]]', '}') -Replace '[\[]', '{'
    ## $a.SaveAsFile("$repopath\$se\$($a.FileName)")
    $a.SaveAsFile("$repopath\$se\$zipName")
    ## $b = Get-Item "$repopath\$se\$($a.FileName)"
    $b = Get-Item "$repopath\$se\$zipName"
    [IO.Compression.ZipFile]::ExtractToDirectory( $b.FullName, "$($b.Directory)\$($b.BaseName)" )
    $fz = $( Get-ChildItem "$($b.Directory)\$($b.BaseName)" -Recurse -File )
    ForEach ( $k in $fz ) {
        $myfname = ($k.Name -Replace '[\]]', '}') -Replace '[\[]', '{'
        $kchksum = Get-CheckSum $k.FullName
        $klmdt = $k.LastWriteTime.ToString('s')
        $cimatch = 0
        ## If ( Test-Path "$repopath\$($k.Name)") { # When the file name already exists, compare the checksum
        If ( Test-Path "$repopath\$($myfname)") { # When the file name already exists, compare the checksum
          ForEach ( $ci in Get-ChildItem "$repopath\$($k.BaseName)*" ) {
                If ( $kchksum -contains $(Get-Checksum $ci) ) { $cimatch ++ }}}
          If ( $cimatch -gt 0 ) { # When there is a match on checksum, don't save the file
              $MG.skippedfiles ++
              ## Add-Content $Logfile -Value "  - **SKIPPING $($k.Name)**"
              Add-Content $Logfile -Value "  - **SKIPPING $myfname**"
              ## Remove-Item "$repopath\$se\$($k.Name)" }
              Remove-Item "$repopath\$se\$myfname" }
          Else { # Move file to destination add try to append a file node
              ## $nom = Get-UnqFilePath $repopath $k.Name
              $nom = Get-UnqFilePath $repopath $myfname
              Move-Item $k.FullName -Destination $nom
              Set-ItemProperty $nom -Name IsReadOnly -Value $true
              $MG.savedfiles ++
              If ( $nom -match '.xls[x|m]?$' ) { $xlfiles += $nom }
              ## If ( $usagemode -eq 'FileOnly' ) { Write-Host "----> $($k.Name) [$($k.Length)_B]" }
              ## Else { $abody += "<li>$($k.Name) [$($k.Length)_B]</li>" }
              ## Add-Content $Logfile -Value "  - **$($k.Name)**, $($k.Length)_Bytes, at $($nom)" }
              If ( $usagemode -eq 'FileOnly' ) { Write-Host -ForeGroundColor Magenta"----> $myfname [$($k.Length)_B]" }
              Else { $abody += "<li>$myfname [$($k.Length)_B]</li>" }
              Add-Content $Logfile -Value "  - **$myfname**, $($k.Length)_Bytes, at $($nom)" }
          $nodefile = $X.SelectSingleNode("/repo/messages/msg[@msgkey='$($msgkey)']/attachments/file[@chksum='$($kchksum)']")
          If ( -not $nodefile ) { # When a file node with the current chksum doesn't exist, append the node
              $banana = Get-Item $nom
              $nodefile = $X.CreateElement('file')
              ## $nodefile.SetAttribute('name',$k.Name)
              $nodefile.SetAttribute('name',$myfname)
              $nodefile.SetAttribute('lmdt',$klmdt)
              $nodefile.SetAttribute('bytes',$k.Length)
              $nodefile.SetAttribute('chksum',$kchksum)
              $nodefile.SetAttribute('href',$([System.Security.SecurityElement]::Escape('file:///' + $nom)))
              $nodeattach.AppendChild($nodefile) | Out-Null }
     } } # End Zip Files
   ElseIf ( $a.FileName -match '^ATT0.*\.htm?$' ) { $MG.sigattachments ++ }
   ElseIf ( $a.FileName -match '^ATT0.*\.txt?$' ) { $MG.sigattachments ++ }
   Else { # Begin Non-Zip Files
    $atfn = $a.FileName -Replace '[\W-[\.]]', '-' 
    ## $a.SaveAsFile( "$repopath\$sc\$($a.FileName)" )
    ## $banana = Get-Item "$repopath\$sc\$($a.FileName)"
    $a.SaveAsFile( "$repopath\$sc\$atfn" )
    $banana = Get-Item "$repopath\$sc\$atfn"
    $bchksum = Get-Checksum $banana.FullName
    $blmdt = $banana.LastWriteTime.ToString('s')
    $cimatch = 0
    ## If ( Test-Path "$repopath\$($a.FileName)" ) { # When the file name already exists, compare the checksum
    If ( Test-Path "$repopath\$atfn" ) { # When the file name already exists, compare the checksum
          ForEach ( $ci in Get-ChildItem "$repopath\$($banana.BaseName)*" ) {
                If ( $bchksum -contains $(Get-Checksum $ci) ) { $cimatch ++ }}}
    If ( $cimatch -gt 0 ) { # When there is a match on checksum, don't save the file
          $MG.skippedfiles ++
          Add-Content $Logfile -Value "  - **SKIPPING $($banana.Name)**"
          Remove-Item $banana }
    Else { # Move file to destination add try to append a file node
          #- $nom = Get-UnqFilePath $repopath $a.FileName
          $nom = Get-UnqFilePath $repopath $banana.Name
          Move-Item $banana -Destination $nom 
          Set-ItemProperty $nom -Name IsReadOnly -Value $true
          #-$lmdt = Get-ChildItem $nom | Select-Object -Prop LastWriteTime
          $MG.savedfiles ++
          If ( $nom -match '.xls[x|m]?$' ) { $xlfiles += $nom }
          ## If ( $usagemode -eq 'FileOnly' ) { Write-Host "----> $($a.FileName) [$($a.Size)_B]" }
          ## Else { $abody += "<li>$($a.FileName) [$($a.Size)_B]</li>" }
          ## Add-Content $Logfile -Value "  - **$($a.FileName)**, $($a.Size)_Bytes, at $($nom)" }
          If ( $usagemode -eq 'FileOnly' ) { Write-Host -ForeGround Magenta "----> $atfn [$($a.Size)_B]" }
          Else { $abody += "<li>$atfn [$($a.Size)_B]</li>" }
          Add-Content $Logfile -Value "  - **$atfn**, $($a.Size)_Bytes, at $($nom)" }
      $nodefile = $X.SelectSingleNode("/repo/messages/msg[@msgkey='$($msgkey)']/attachments/file[@chksum='$($bchksum)']")
      If ( -not $nodefile ) { # When a file node with the current chksum doesn't exist, append the node
          $nodefile = $X.CreateElement('file')
          #- $nodefile.SetAttribute('name',$a.FileName)
          #- $nodefile.SetAttribute('lmdt',$lmdt.LastWriteTime.ToString('s'))
          #- $nodefile.SetAttribute('bytes',$a.Size)
          $nodefile.SetAttribute('name',$banana.Name)
          $nodefile.SetAttribute('lmdt',$blmdt)
          $nodefile.SetAttribute('bytes',$banana.Length)
          $nodefile.SetAttribute('chksum',$bchksum)
          $nodefile.SetAttribute('href','file:///' + $([uri]::EscapeDataString($nom)))
          $nodeattach.AppendChild($nodefile)  | Out-Null }
      } # End Non-Zip Files
    }
  }


# Apply the message to the SharePoint List
If ( $nodemsg ) { # Begin SP List Item
# Attempt to select from the SP list on the message conversation ID
$xmlImage = New-Object System.Xml.XmlDocument
$elemntQuery = $xmlImage.CreateElement("Query")
$elemntViewFld  = $xmlImage.CreateElement("ViewFields")
$elemntQueryOpt = $xmlImage.CreateElement("QueryOptions")
$elemntQuery.InnerXML = @"
<Where>
  <Contains>
     <FieldRef Name="MessageLink" />
     <Value Type="Text">$($P.linksrv + $msgkey)</Value>
  </Contains>
</Where>
"@
$elemntViewFld.InnerXML = @"
<FieldRef Name="ID" />
<FieldRef Name="MessageLink" />
"@
$elemntQueryOpt.InnerXML = @"
<IncludeMandatoryColumns>FALSE</IncludeMandatoryColumns>
<DateInUtc>TRUE</DateInUtc>
"@
$ndReturnSEL = $objSP.GetListItems($P.SPListGUID, $null, $elemntQuery, $elemntViewFld, $rowLimit, $elemntQueryOpt, $null)

# Apply the result to the SP List (update or insert)
If ($ndReturnSEL.Data.ItemCount -gt 0 ) {
$MG | Add-Member –M NoteProperty –Name sp –Val   'skip-duplicate'
Write-Warning "MessageLink is already populated in SharePoint"
} 
Else {
# Insert a PS List Item
If ( $xlfiles.Count -gt 0 ) { #Check for guess values 
    If ( -not $Excel ) { $Excel = New-Object -ComObject Excel.Application }
    ForEach ( $xl in $xlfiles ) {
        $tmpEAP = $ErrorActionPreference
        $ErrorActionPreference = 'SilentlyContinue'
        $wb = $Excel.Workbooks.Open($xl)
        $ErrorActionPreference = $tmpEAP
        If ( $wb ) {
          # lotus
          $properties = $wb.BuiltInDocumentProperties
          ForEach($property in $properties) {
          $pn = [System.__ComObject].InvokeMember("name",$binding::GetProperty,$null,$property,$null)
          If ( $pn -contains 'Content status' ) { 
              $foundstat =  [System.__ComObject].invokemember("value",$binding::GetProperty,$null,$property,$null) } }
               If ($foundstat -Match "v[0-9]+") { 
                   If ( $versions -contains $Matches[0].ToUpper() ) { 
                   $wb.Close($false)
                  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null
                  Remove-Variable wb
                  [System.GC]::Collect()
                  $strDataSrc = "Data Source = $xl"
                  $objFormConn    = New-Object System.Data.OleDb.OleDbConnection("$($P.OLEDBProvider);$strDataSrc;$($P.OLEDBExtend)")
                  $objFormConn.Open()

                  $sqlCommand = New-Object System.Data.OleDb.OleDbCommand($NIQuery)
                  $sqlCommand.Connection  = $objFormConn

                  $objAdapter = New-Object "System.Data.OleDb.OleDbDataAdapter"
                  $objAdapter.SelectCommand = $sqlCommand

                  $DataTable  = New-Object "System.Data.DataTable"
                  $feedback   = $objAdapter.Fill($DataTable)

                  $sqlCommand.Dispose()
                  $objFormConn.Close()
                  $objFormConn.Dispose()

                  If ( $feedback -gt 0 ) {
                      $foundcat = $DataTable.Rows[0].hCategory
                      $foundbrd = $DataTable.Rows[0].Brand }
                   } }
              # l-m 
               Else {
                  $ws = $Excel.Sheets | Where-Object { $_.Name -contains 'New Item Setup Form' }
                  If ( $ws ) {
                        $foundcat = $ws.Cells.Item(23,4).Value2
                        If ( $foundcat ) { $categoryguess = $foundcat }
                        $foundbrd = $ws.Cells.Item(23,9).Value2
                        If ( $foundbrd ) { $brandguess = $foundbrd } }
                  $wb.Close($false)
                  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null
                  Remove-Variable wb
                  [System.GC]::Collect() } }
        Else { Write-Warning "Excel failed to OPEN for reading guesses: $xl"}      
      }
}
$escapemsgkey = [System.Security.SecurityElement]::Escape($msgkey)
$strML = @"
<Method ID='1' Cmd='New'>
<Field Name='Title'>$($([System.Security.SecurityElement]::Escape($i.ConversationTopic)))</Field>
<Field Name='Subject'>$($([System.Security.SecurityElement]::Escape($i.Subject)))</Field>
<Field Name='BrandName'>$($([System.Security.SecurityElement]::Escape($brandguess)))</Field>
<Field Name='ReviewCategory'>$($([System.Security.SecurityElement]::Escape($categoryguess)))</Field>
<Field Name='ReviewRound'>$($rnd)</Field>
<Field Name='ContactEmail'>$($sender)</Field>
<Field Name='ReviewStatus'>Acknowledged</Field>
<Field Name='PendingAction'>None</Field>
<Field Name='MessageLink'>$($P.linksrv + $($escapemsgkey))</Field>
</Method>
"@

# <Field Name='BrandName'>$($brandguess)</Field>
# <Field Name='ReviewCategory'>$($categoryguess)</Field>

$xmlImage = New-Object System.Xml.XmlDocument
$ele = $xmlImage.CreateElement("Batch")
$ele.SetAttribute("OnError","Continue")
$ele.SetAttribute("ListVersion","1")
$ele.SetAttribute("ViewName",$P.SPViewGUID)
$ele.InnerXML = $strML

$ndReturn = $objSP.UpdateListItems($P.SPListGUID, $ele)
# $ndReturn.Result | Select-Object -Property @{l='iD';e={$_.ID[0]}},ErrorCode,row
# If ( $ndReturn.Result.ErrorCode -match '0x00000000' ) { Write-Host "----+ SP Appended"}
$MG | Add-Member –M NoteProperty –Name sp –Val 'insert-listitem'

} }  # End SP List Item
$ndReturnSEL = $null

$moved += $i.Move($oltrgfolder) | Out-Null

If ( $usagemode -ne 'FileOnly' ) {
$rail = $i.Reply()
$rail.SentOnBehalfOfName = $P.mailbox
$abody += $P.closing
$rail.HTMLBody = $abody
If ( $usagemode -eq 'MsgDisplay' ) { $rail.Display() }
ElseIf ( $usagemode -eq 'MsgSend' ) { $rail.Send() } }

Write-Host -Fore Yellow "Summary for [ $msgkey ]"
$MG | Format-List *
If ( $compareimage ) { $compareimage.Delete() ; Remove-Variable compareimage }
$fsimage.Delete()
$card++
<# End Message Loop
 ___       __            ___  __   __        __   ___          __   __   __
|__  |\ | |  \     |\/| |__  /__` /__`  /\  / _` |__     |    /  \ /  \ |__)
|___ | \| |__/     |  | |___ .__/ .__/ /~~\ \__> |___    |___ \__/ \__/ |
#>
}
$X.Save($P.message_repo)
Write-Host "    Messages  Flagged: $($colmsg.Count)"
Write-Host "=== Messages Exported: $($moved.Count)"
Add-Content $Logfile -Value "$nl#####Message Processing Complete [$($moved.Count)]$nl"

<#  Wrapup
      __        __        __
|  | |__)  /\  |__) |  | |__)
|/\| |  \ /~~\ |    \__/ |
#>
# Get-Variable
If ( Test-Path("$repopath\$sc")  ) { Remove-Item -Path "$repopath\$sc*" -Recurse}
If ( Test-Path("$repopath\$se")  ) { Remove-Item -Path "$repopath\$se*" -Recurse}
$X = $null
$objSP.Dispose()
$objSP = $null
$xmlImage = $null
$rail = $null
$moved = $null
$colmsg = $null
$olsrcfolder = $null
$oltrgfolder = $null
$olApp = $null
### $olApp.Quit()
###[System.Runtime.Interopservices.Marshal]::ReleaseComObject($olApp) | Out-Null
If ( $Excel ) { $Excel.Quit()
      [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
      Remove-Variable Excel
      [System.GC]::Collect() }
If ($fsimage) {Remove-Variable fsimage}
Remove-Variable olApp
[System.GC]::Collect()

#### KTHXBYE ####
