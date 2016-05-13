<#
# BreakPoint Syntax #
Set-PSBreakPoint -Script hadar_byc.ps1 -Variable foo
Set-PSBreakpoint -Script hadar_byc.ps1 -Command Send-MailMessage
Remove-PSBreakPoint -ID 3
Get-PSBreakPoint | Remove-PSBreakPoint
Enable-PSBreakPoint -ID 3
Disable-PSBreakPoint -ID 3
#>
## Functions

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
		$checksum	+= $byte.ToString('x2')
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
ForEach ($byte in $bash) { $checksum	+= $byte.ToString('x2') }
$checksum
}
Function Get-UnqXlsPath ( $pstr, $nstr) {
If ( Test-Path "$($pstr)\$($nstr)" ) {
$arr = $nstr.Split([Char]46)
$s1 = $arr[0..($arr.length - 2)]
$cx = $( Get-ChildItem "$($pstr)\$($s1 -Join [Char]46)*" ).Name
$s1 += Get-SomeCheck $cx
$s1 += $arr[-1]
"$($pstr)\$($s1 -Join [Char]46)" }
Else  { "$($pstr)\$($nstr)" }  }

# Begin Setup
# $usagemode ValidateSet('Debug','Test','Normal')

Set-Location c:\psg
$usagemode = 'Debug'
$parmfile  = 'Hadar_Parms_JI'
$P = Get-Content ".\$parmfile.pson" | Out-String | Invoke-Expression
$apath = $P.folder_stage
$Logfile = $([System.Environment]::GetEnvironmentVariable('TMP','MACHINE')) + "\$($P.logfile_name).md"
If (!(Test-Path $Logfile)) {
   Try { Set-Content -Path $Logfile -Value ($null) }
   Catch { Write-Error "Logfile not found/valid: $($Logfile)"; Break } }
$nl = [char]13
$sc = '~sample~'
$se = '~extrct~'
$testTo = ''
If ( $usagemode -eq 'Test') {
      $testTo = $($env:UserName + '@wholefoods.com') }
[Reflection.Assembly]::LoadWithPartialName('System.IO.Compression.FileSystem') | Out-Null

## End Setup

## Connect to Outlook

Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"
$olApp     = New-Object -ComObject Outlook.Application
$namespace = $olApp.GetNameSpace("MAPI")

## Get Outlook Folders

$parchive = $( Recurse-Folders $namespace $P.inbox_trg )[0]
$parchive.FolderPath
$pdiscard = $( Recurse-Folders $namespace $P.inbox_push )[0]
$pdiscard.FolderPath
$ppndg = $( Recurse-Folders $namespace $P.inbox_src )[0]
$ppndg.FolderPath

## A03 Begin Sort

$messages = @()
$discards = @()
$msgfiles = 0
$totfiles = 0
Do {
ForEach ( $m in  $ppndg.Items  ) {
$mtest = $false
If     ( $($m.Attachments).FileName -match '.xls[x|m]?$' ) { $mtest = $true }
ElseIf ( $($m.Attachments).FileName -match '.zip?$' ) {
ForEach ( $f in $m.Attachments ) {
If ( $f.FileName -match '.zip' ) {
$f.SaveAsFile("$apath\$sc$($f.FileName)")
If ( $([IO.Compression.ZipFile]::OpenRead("$apath\$sc$($f.FileName)").Entries).Name -match '.xls[x|m]?$' ) { $mtest = $true }
Else { $skipped += 1 }
}  }  }
Else { $skipped += 1 }
If ( $mtest ) { $messages += $m.Move( $parchive ) }
Else { $discards += $m.Move( $pdiscard ) }
}
$m.subject
} While ( $ppndg.Items.Count -gt 0 )
Remove-Item "$apath\$sc*.zip"

## A03 End Sort

<#
$messages.count
$($messages[0].Sender.Address).Split([Char]61)[-1] + '@wholefoods.com'

If ( $usagemode -eq 'Test' ) { $ReplyTo = $($env:UserName + '@wholefoods.com') }
ElseIf ( $*.Sender.Type -eq 'EX' ) {
$ReplyTo = $($*.Sender.Address).Split([Char]61)[-1] + '@wholefoods.com' }
Else { $ReplyTo = $*.Sender.Address }




$messages[0] | gm
$messages[0].SenderEmailType
$discards[0].SenderEmailAddress
$messages[0].Attachments
$messages[0].Recipients
$discards[0].Display
$messages[0].Sender.Address
$discards[0].Sender.Address
$messages[0].Sender.PropertyAccessor
$messages[0].From
$($messages[5].Attachments).FileName
Send-MailMessage -To "Jeff Inman (CE CEN)" `
      -From $($env:UserName + '@wholefoods.com') `
      -Subject "Test" `
      -Body "Hello World!" `
      -SMTPServer smtp.wfm.pvt
$nom = $($messages[5].Attachments)[0].FileName
$($messages[5].Attachments)[0].SaveAsFile("$apath\$nom")
Get-ChildItem $apath
Get-SomeCheck $messages[0].EntryID
#>

<# Begin Simple Staging

ForEach ( $m in $messages ) {
      ForEach ( $a in $m.Attachments ) {
            If ( $a.FileName -match '.xls[x|m]?$' ) {
                 $a.SaveAsFile("$apath\$($a.FileName)")
               }
            ElseIf ( $a.FileName -match '.zip?$' ) {
                 $a.SaveAsFile("$apath\$se$($a.FileName)")
               }
   }     }

## End Simple Staging
#>

## Begin zip extract Staging

$reply = @{}
ForEach ( $m in $messages ) {
$mn = "_$( Get-SomeCheck $m.EntryID )"
$reply += @{ $mn = @{ 'sender' = $m.Sender ; 'subject' = $m.Subject ; 'xlfiles' = @{} } }
ForEach ( $a in $m.Attachments ) {
If ( $a.FileName -match '.xls[x|m]?$' ) {
$nom = Get-UnqXlsPath $apath $a.FileName
$fn = "_$( Get-SomeCheck $nom )"
$a.SaveAsFile( $nom )
$reply.$($mn).xlfiles += @{ $fn = @{ 'name' = $a.FileName ; 'size' = $a.Size ; 'full' = $nom } }
}
ElseIf ( $a.FileName -match '.zip?$' ) {
$a.SaveAsFile("$apath\$se$($a.FileName)")
$b = Get-Item "$apath\$se$($a.FileName)"
[IO.Compression.ZipFile]::ExtractToDirectory( $b.FullName, "$($b.Directory)\$($b.BaseName)" )
$fz = $( Get-ChildItem "$($b.Directory)\$($b.BaseName)" -Recurse -File | Where-Object {$_.Name -match ".xls[x|m]?$"} )
ForEach ( $k in $fz ) {
$nom = Get-UnqXlsPath $apath $k.Name
$fn = "_$( Get-SomeCheck $nom )"
$reply.$($mn).xlfiles += @{ $fn = @{ 'name' = $k.Name ; 'size' = $k.Length ; 'full' = $nom } }
Move-Item $k.FullName -Destination $nom
}   }   }
$m.Subject }
Remove-Item "$apath\$se*" -Recurse

<#
$messages.count
$reply.count
#>

## Continuing zip extract Staging ...
## ... with messaging
<#
If ( $usagemode -ne 'Debug' ) {
ForEach ( $m in $reply.Keys ) {
      $abody = $P.opening
      ForEach ( $n in $reply.$($m).xlfiles.Keys ) {
             $abody += "<li>$($reply.$($m).xlfiles.$($n).name) [$($reply.$($m).xlfiles.$($n).size)B]</li>"
         }
      $abody += $P.closing
If ( $usagemode -eq 'Test' ) { $ReplyTo = $($env:UserName + '@wholefoods.com') }
ElseIf ( $*.Sender.Type -eq 'EX' ) {
$ReplyTo = $($*.Sender.Address).Split([Char]61)[-1] + '@wholefoods.com' }
Else { $ReplyTo = $*.Sender.Address }
      Send-MailMessage -To $P.ccemail `
             -From $($env:UserName + '@wholefoods.com') `
             -Subject $($P.prepend + $reply.$($m).subject) `
             -Body $abody -BodyAsHtml `
             -SMTPServer smtp.wfm.pvt
   }   }
#>
## End zip extract Staging

<# Iterate Over $reply
ForEach ( $m in $reply.Keys ) {
      $reply.$($m).subject
      ForEach ( $n in $reply.$($m).xlfiles.Keys ) {
            Get-Checksum $reply.$($m).xlfiles.$($n).full
   }     }
.
ForEach ( $m in $reply.Keys ) {
      $abody = $P.salutation
      ForEach ( $n in $reply.$($m).xlfiles.Keys ) {
             $abody += "<li>$($reply.$($m).xlfiles.$($n).name) $($reply.$($m).xlfiles.$($n).size)B</li>"
         }
$abody += $P.signature
$reply.$($m).subject
$P.ccemail
$abody
}

#>


<#

Remove-Item "$apath\*"

$ffset += @( $i.Name, $i.Length, 'k' )


Get-SomeCheck $(Get-ChildItem c:\tmp).Name

Get-ChildItem "$apath\$($se)*.zip"

[IO.Compression.ZipFile]::ExtractToDirectory("$apath\$($se)sum.zip", "$apath\$($se)sum")


Get-ChildItem "$apath\$($se)sum" -Recurse -File |
Where-Object {$_.Name -match ".xls[x|m]?$"} |
Tee-Object -Variable zx |
Move-Item -Destination $apath

$zx
#>

# Tee-Object -Variable <String> [-InputObject <PSObject> ] [ <CommonParameters>]
<#



$foundxls += @{
"id$($bar)" = @{ 'name' = $a.FileName ; 'size' = $a.Size } }
$bar += 1


$foundxls = @{}
ForEach ( $m in $messages ) {
$foundxls += @{ "EI$($m.EntryID)" = @{
'sender'  = $m.SenderName
'subject' = $m.Subject
'attcount' = $m.Attachments.Count}   }
}

$Attachments = @{
      'i101' = @{
            'sender' = 'joe@foo.com'
            'subject' = 'form submission'
            'xlf' = @{
                  'name' = 'baz.xlsx' ; size = 99
               }
         }
   }

$Attachments = @{
      'i101' = @{
            'sender' = 'joe@foo.com'
            'subject' = 'form submission'
            'xlf' = @{
                  'name' = 'baz.xlsx' ; size = 99
                     }
            'zif' = @{
                  'sum.zip' = @{
                        'name' = 'lorem.xlsx'  ; size = 111
                     }
               }
         }
   }
.
$Attachments = @{}
$Attachments += @{ 'i101' = @{ 'sender' = 'joe@foo.com' ; 'subject' = 'form submission' ; 'xlf' = @{} ; 'zif' = @{} ; 'otf' = @{} }}
$Attachments += @{ 'i301' = @{
'sender' = 'joe@foo.com'
'subject' = 'form submission'
'xlf' = @{} ; 'zif' = @{} ; 'otf' = @{} }}


$($($ppndg.Items[0]).Attachments).FileName -match '.zip'

$ppndg.Items.Count
$($($ppndg.Items[0]).Attachments).FileName
$nom = $($($ppndg.Items[0]).Attachments)[1].FileName
$($($ppndg.Items[0]).Attachments)[1].SaveAsFile("$apath\$nom")

SaveAsFile( "$apath\$($m.Subject) $($a.Filename)" )

$zpath = "$apath\" + [guid]::NewGuid()
New-Item -ItemType Directory $zpath | Out-Null
Remove-Item $zpath
[Reflection.Assembly]::LoadWithPartialName('System.IO.Compression.FileSystem') | Out-Null
$([IO.Compression.ZipFile]::OpenRead("$apath\$nom").Entries).Name

[IO.Compression.ZipFile]::ExtractToDirectory(String, $zpath)

Get-ChildItem c:\psg\util\sum.zip
[IO.Compression.ZipFile]::ExtractToDirectory('c:\psg\util\sum.zip', 'c:\psg\util')
Remove-Item c:\psg\util\ipsum

$zpath = "$apath\" + [guid]::NewGuid()
New-Item -ItemType Directory $zpath | Out-Null
$nom = $($($ppndg.Items[0]).Attachments)[1].FileName
$($($ppndg.Items[0]).Attachments)[1].SaveAsFile("$zpath\$nom")
$([IO.Compression.ZipFile]::OpenRead("$zpath\$nom").Entries).Name
[IO.Compression.ZipFile]::ExtractToDirectory("$zpath\$nom", $zpath)
ForEach ( $i in $(Get-ChildItem $zpath -Recurse -File '*.xls?') ) { $i.Name }

ForEach ( $m in $($ppndg.Items | Where-Object { $_.Class -eq 43 } ) {
ForEach ( $a in $m.Attachments ) {
$fcount = 0
If     ( $($([IO.Compression.ZipFile]::OpenRead("$zpath\$nom").Entries).Name -like '*.xls?').Count > 0 ) {}
ElseIf ( $($([IO.Compression.ZipFile]::OpenRead("$zpath\$nom").Entries).Name -like '*.zip').Count  > 0 ) {}
}
}

$([IO.Compression.ZipFile]::OpenRead("$apath\$($se)sum.zip").Entries).Name

[IO.Compression.ZipFile]::ExtractToDirectory("$apath\$($se)sum.zip", "$apath\$($se)sum")

Get-ChildItem "$apath\$($se)sum" -Recurse -File | Where-Object {$_.Name -match ".xls[x|m]?$"}


.
ForEach
Read/Count
Move/Stage


$([IO.Compression.ZipFile]::OpenRead("$zpath\$nom").Entries).Name -contains '*.xls?'

$($([IO.Compression.ZipFile]::OpenRead("$zpath\$nom").Entries).Name -like '*.xls?').Count

## A03 End
## A01 Begin

Do {
      ForEach ( $m in $($ppndg.Items | Where-Object { $_.Class -eq 43 } ) )  {
            ForEach ( $a in $m.Attachments ) {
                $a.SaveasFile( "$apath\$($m.Subject) $($a.Filename)" )
           } $messages += $m.Move( $parchive )
         }
   } While ( $ppndg.Items.Count -gt 0 )

## Post Process - List saved files and send receipt message

ForEach ( $m in $messages ) {
      Add-Content $Logfile -Value "-  Message[ $($m.Subject) ] Attachment Count: $($m.Attachments.Count)"
      ForEach ( $a in $m.Attachments ) {
          $abody += "<li>$($a.FileName)</li>"
            $totfiles += 1
            $afullpath = "$apath\$($m.Subject) $($a.FileName)"
            If (! $(Test-Path $afullpath )) { Write-Output "Attachment Not Saved??!!?" }
            Else { $aobj = Get-ChildItem $afullpath
            Add-Content $Logfile -Value "  - **$($aobj.Name)**, $($aobj.Length)_Bytes" }
         }
     If ( $usagemode -ne 'Debug' ) {
           $mbody  = $replybodybegin
         $mbody += $abody
         $mbody += $replybodyend
         If ( $usagemode -eq 'Test') { $mTo = $testTo }
         Else { $mTo = $m.From }
         Send-MailMessage -To $mTo `
             -From $($env:UserName + '@wholefoods.com') `
             -Subject $("RE: " + $m.Subject) `
             -Body $mbody -BodyAsHtml `
             -SMTPServer smtp.wfm.pvt
      }
      Clear-Variable abody
   }
## A01 End

#>
