<################### budgie #####################
         _               _       _      
        | |__  _   _  __| | __ _(_) ___ 
        | '_ \| | | |/ _` |/ _` | |/ _ \
        | |_) | |_| | (_| | (_| | |  __/
        |_.__/ \__,_|\__,_|\__, |_|\___|
                           |___/        

 Load 2017 EDL Forms into an MS Access Database

#################################################

.\budgie_2017.ps1 Budgie_Parms_G

#>
<## To Do ####################################

log step activity 

log exception messages    Get-ChildItem $P.ffpath

##############################################>


Param
 (
 # Name of local parameter file
 [Parameter(Mandatory=$true, Position=0)]
 [ValidateNotNull()] [ValidateNotNullOrEmpty()]
 $parmfile )

# $parmfile = 'Budgie_Parms_G'

If ( -not $(Test-Path "c:\psg\$parmfile.pson") ) {
      Write-Host -BackgroundColor Magenta -ForegroundColor Black `
      "!! Parameter File not found: c:\psg\$parmfile.pson !!"
      Break }
Try {
      $T = Get-Content "c:\psg\$parmfile.pson" | Out-String | Invoke-Expression }
Catch {
      Write-Host -BackgroundColor Magenta -ForegroundColor Black `
      "!! Parameter File Failed: c:\psg\$parmfile.pson !!"
      Break }

$P = New-Object �TypeName PSObject
$P | Add-Member �M NoteProperty �Name edlyear �Val $T.edlyear
$P | Add-Member �M NoteProperty �Name bpath   �Val $T.bpath
$P | Add-Member �M NoteProperty �Name dbname  �Val $T.dbname
$P | Add-Member �M NoteProperty �Name tblf    �Val $T.tblf
$P | Add-Member �M NoteProperty �Name tbli    �Val $T.tbli
$P | Add-Member �M NoteProperty �Name fzip    �Val $T.fzip
$P | Add-Member �M NoteProperty �Name ffpath  �Val $($P.bpath + [char]92 + 'forms' )
$P | Add-Member �M NoteProperty �Name strdb   -Val $($P.bpath + [char]92 + $P.dbname )
# "Provider=Microsoft.ACE.OLEDB.12.0"
# $strExtend  = "Extended Properties=Excel 12.0"   ## Consider usage for HDR and IMEX ( header flag & import/export mode )


<#
$P = New-Object �TypeName PSObject
$P | Add-Member �M NoteProperty �Name edlyear �Val '2016'
$P | Add-Member �M NoteProperty �Name bpath   �Val 'c:\tmp\budgie\WB\APR26' 
$P | Add-Member �M NoteProperty �Name dbname  �Val 'WB_EDL_asof_MAY11.accdb'
$P | Add-Member �M NoteProperty �Name tblf    �Val 'EDL_Form'
$P | Add-Member �M NoteProperty �Name tbli    �Val 'EDL_Item'
$P | Add-Member �M NoteProperty �Name fzip    �Val 'budgie_in2016WBAPR26.zip'
$P | Add-Member �M NoteProperty �Name ffpath  �Val $($P.bpath + [char]92 + 'forms' )
$P | Add-Member �M NoteProperty �Name strdb   -Val $($P.bpath + [char]92 + $P.dbname )
#>
<#
$P = New-Object �TypeName PSObject
$P | Add-Member �M NoteProperty �Name edlyear �Val '2016'
$P | Add-Member �M NoteProperty �Name bpath   �Val 'C:\tmp\budgie\G\Shrink_Allowance\APR08'
$P | Add-Member �M NoteProperty �Name dbname  �Val '2016_EDL_Database_GSA_asofAPR27.accdb'
$P | Add-Member �M NoteProperty �Name tblf    �Val 'EDL_Form'
$P | Add-Member �M NoteProperty �Name tbli    �Val 'EDL_Item'
$P | Add-Member �M NoteProperty �Name fzip    �Val 'budgie_in2016APR08.zip'
$P | Add-Member �M NoteProperty �Name ffpath  �Val $($P.bpath + [char]92 + 'forms' )
$P | Add-Member �M NoteProperty �Name strdb   -Val $($P.bpath + [char]92 + $P.dbname )
#>

Set-Location $P.ffpath
If (Test-Path "$($P.bpath)\$($P.fzip)") { Move-Item "$($P.bpath)\$($P.fzip)" "$($P.bpath)\bkup_$($P.fzip)" -Force }
Get-ChildItem $P.ffpath | Write-Zip -OutputPath "$($P.bpath)\$($P.fzip)"

<# Zip File Syntax
$fpath = '\\wfm-team\team\RegionalPurchasing\National Promotions\2016 EDLC Folders\2016 Global Grocery EDLC Files\2016 EDLC Submissions'
$repo = 'c:\tmp\budgie\G\SEP29\budgie_in2016SEP29.zip'
$target = 'c:\tmp\budgie\G\AUG14\forms'
# Remove-Item $target
Set-Location $fpath
Set-Location $P.bpath
Get-ChildItem $fpath -Recurse | Write-Zip -OutputPath $($P.bpath + '\' + $P.fzip)
Get-ChildItem $P.ffpath -Recurse | Write-Zip -OutputPath $($P.bpath + '\' + $P.fzip)
[Reflection.Assembly]::LoadWithPartialName('System.IO.Compression.FileSystem') | Out-Null
[IO.Compression.ZipFile]::ExtractToDirectory( $repo, $target )
[IO.Compression.ZipFile]::ExtractToDirectory( $($P.bpath + [char]92 + $P.fzip ), $P.ffpath )
Set-Location 'c:\psg'

$xpath = 'C:\tmp\budgie\G\SEP29\pptx_summaries'
$tpath = '\\wfm-team\team\RegionalPurchasing\National Promotions\2016 EDLC Folders\2016 Global Grocery EDLC Files\SEP29\2016_PPTX_Summaries_asofSEP29.zip'
Set-Location $xpath
Get-ChildItem $xpath -Recurse | Write-Zip -OutputPath $tpath

#>
<# Misc Syntax            $P.ffpath
# $traceability.Length

[char]92

Get-ChildItem $strRepoFile  "C:\tmp\banana.accdb" 'c:\tmp\budgie\temp'

Get-ChildItem 'c:\tmp\budgie\temp'
Get-Item "c:\tmp\budgie\temp\2015-2016 WFM EDLC Yogi Teas.xlsm" | fl *
$baz = Get-Item "c:\tmp\budgie\temp\2015-2016 WFM EDLC Annie's.xlsm"
$baz = Get-Item 'c:\tmp\budgie\temp\2015-2016 WFM EDLC Yogi Teas.xlsm'

$baz.Name -replace [Char]39, ([Char]92 + [Char]39)
$baz | fl *


Name
LastWriteTimeUTC
--CheckSum--

$in5file = @"
Insert Into zloadwork (Traceability, EDL_Year, File_Name, File_Butes, Last_Modified, MD5_Checksum)
Values ( '$($traceability)', '$($edlyear)', '$($baz.Name)', '$($baz.LastWriteTimeUTC)', '$(Get-Checksum $baz.FullName)')
"@

#>

Function Test-UPCaHasCheckDigit { param ([string]$upc)
If ($upc.Length -ne 13) { Write-Warning "UPC ($upc) length is not thirteen characters!" ; Return 0 }
$k = 0
ForEach ($i in $upc[0,2,4,6,8,10])  { $k +=   ([int]$i.ToString()) }
ForEach ($i in $upc[ 1,3,5,7,9,11]) { $k += 3*([int]$i.ToString()) }
(10 - ($k % 10) -eq [int]$upc[12].ToString() )
}

Function Get-Checksum ($file, $crypto_provider) {
  If ($crypto_provider -eq $null) {
    $crypto_provider = New-Object -TypeName System.Security.Cryptography.MD5CryptoServiceProvider
   }
  $file_info = Get-Item $file
  Trap { Continue }
  $stream = $file_info.OpenRead()
  If ($? -eq $false) {
      Return $null
   }
  $bytes    = $crypto_provider.ComputeHash($stream)
  $checksum = ''
  ForEach ($byte in $bytes) {
      $checksum    += $byte.ToString('x2')
   }
  $stream.Close() | Out-Null
  Return $checksum

 }

$P | Add-Member �M NoteProperty �Name traceability �Val $(Get-Checksum ($P.bpath + [char]92 + $P.fzip))
# $strRepoSrc  = "Data Source = $($P.strdb)"

<#  $P.traceability
[Reflection.Assembly]::LoadWithPartialName('System.IO.Compression.FileSystem') | Out-Null
[IO.Compression.ZipFile]::ExtractToDirectory( 'c:\tmp\budgie\budgie_in2016.zip', 'c:\tmp\budgie\temp' )
#>

<##############################################################

 _              _         __                 _ _    _   
| |__  ___ __ _(_)_ _    / _|___ _ _ _ __   | (_)__| |_ 
| '_ \/ -_) _` | | ' \  |  _/ _ \ '_| '  \  | | (_-<  _|
|_.__/\___\__, |_|_||_| |_| \___/_| |_|_|_| |_|_/__/\__|
          |___/                                         

##############################################################>

## Open an MS Excel Image
$Excel = New-Object -ComObject Excel.Application

## Open an MS Access as an OLEDB Source
$strRPrvidr  = "Provider=Microsoft.ACE.OLEDB.12.0"
$strRepoSrc  = "Data Source = $($P.strdb)"
$objRepoConn = New-Object System.Data.OleDb.OleDbConnection("$strRPrvidr;$strRepoSrc")
$objRepoConn.Open()

$batch = Get-ChildItem $P.ffpath
$n = 1
$swf = [Diagnostics.Stopwatch]::StartNew()
ForEach ($i in $batch) {

Write-Progress -Activity "Building Form List" -Status $i.Name -PercentComplete $n
$fmd5 = Get-Checksum $i.FullName
$Wrkbk = $Excel.Workbooks.Open($i.FullName)
$fbrand = $Excel.Range('ibrand').Value2 -replace [Char]39, ([Char]39 + [Char]39)
# $fbrand = $Excel.Range('ibrand').Text -replace [Char]39, ([Char]39 + [Char]39)
If ( $fbrand.Length -le 1 ) { $fbrand = 'mt' }


$in5file = @"
Insert Into $($P.tblf) (Traceability, EDL_Year, File_Name, File_Bytes
, Last_Modified, MD5_Checksum
, Brand_Name, Brand_Contact_1_Name, Brand_Contact_1_Email
, Brand_Contact_2_Name, Brand_Contact_2_Email
, Broker_Contact_Name, Broker_Contact_Email)
Values ( '$($P.traceability)', '$($P.edlyear)', '$($i.Name -replace [Char]39, ([Char]39 + [Char]39))', '$($i.Length)'
, '$($i.LastWriteTimeUTC)', '$($fmd5)'
, '$($fbrand)', '$($Excel.Range('iAuthName').Value2 -replace [Char]39, ([Char]39 + [Char]39))', '$($Excel.Range('iAuthEmail').Value2)'
, '$($Excel.Range('iBillingName').Value2 -replace [Char]39, ([Char]39 + [Char]39))', '$($Excel.Range('iBillingEmail').Value2)'
, '$($Excel.Range('iBrokerName').Value2 -replace [Char]39, ([Char]39 + [Char]39))', '$($Excel.Range('iBrokerEmail').Value2)'
)
"@


# $in5file
$sqlFilecmd = New-Object System.Data.OleDb.OleDbCommand($in5file)
$sqlFilecmd.Connection  = $objRepoConn
$sqlFilecmd.CommandText = $in5file
$feedback = $sqlFilecmd.ExecuteNonQuery()
# $feedback
$sqlFilecmd.Dispose()

$Wrkbk.Close($false)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Wrkbk) | Out-Null
Remove-Variable Wrkbk
[System.GC]::Collect()
$n++
}

$swf.Stop()
$swf.Elapsed | Select-Object -Property Minutes,Seconds | Format-Table -AutoSize

# Close Excel Application
$Excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
Remove-Variable Excel
[System.GC]::Collect()

# $objRepoConn.Close()

<##############################################################

             _    __                 _ _    _   
 ___ _ _  __| |  / _|___ _ _ _ __   | (_)__| |_ 
/ -_) ' \/ _` | |  _/ _ \ '_| '  \  | | (_-<  _|
\___|_||_\__,_| |_| \___/_| |_|_|_| |_|_/__/\__|


##############################################################>

### Stage Array of Form/File Names

<# Open an MS Access as an OLEDB Source
$strRPrvidr  = "Provider=Microsoft.ACE.OLEDB.12.0"
$strRepoSrc  = "Data Source = $($P.strdb)"
$objRepoConn = New-Object System.Data.OleDb.OleDbConnection("$strRPrvidr;$strRepoSrc")
#>
# $objRepoConn.Open()

$se2formlist = @"
Select File_name From $($p.tblf)
"@

$sqlFilecmd = New-Object System.Data.OleDb.OleDbCommand($se2formlist)
$sqlFilecmd.Connection  = $objRepoConn
$sqlFilecmd.CommandText = $se2formlist

$aryWB = @()
$DataReader = $sqlFilecmd.ExecuteReader()
While($DataReader.Read()) {  $aryWB += $DataReader[0].Tostring()  }
$DataReader.Close()

# $aryWB = @('2015-2016 WFM EDLC Annie Chuns.xlsm','2015-2016 WFM EDLC Blue Diamond HP Hood.xlsm')
# $strFileNam = $bbase + '\temp\' + $aryWB[0]
# Clear-Host
# $aryWB.Count

<# Old SQL
$strQuery   = @"
Select [cYear], [cFr_Date], [cTo_Date], [cFamily]
, [cBrand], [cItem_Desc], [cCase_Pk], [cSz], [cUoM], [cUPC]
, [cList_Cost], [cCase_Freight], [cFOB_Flag]
, [cProgram_Type], [cPct_Off], [cSB_AMT], [cREG_SRP], [cEDL_SRP]
From [EDLC Agreement Detail$]
Where [cUPC] is not NULL
And [cSz] <> -37
And [cItem_Desc] not like 'EXAMPLE -%'
"@
#>

<##############################################################

 _              _        _ _               _              _ 
| |__  ___ __ _(_)_ _   (_) |_ ___ _ __   | |___  __ _ __| |
| '_ \/ -_) _` | | ' \  | |  _/ -_) '  \  | / _ \/ _` / _` |
|_.__/\___\__, |_|_||_| |_|\__\___|_|_|_| |_\___/\__,_\__,_|
          |___/                                             
          
##############################################################>

ForEach ( $j in $aryWB ) {

## Open an MS EXCEL Source
## $strFileNam = $P.ffpath + [char]92 + '2015-2016 WFM EDLC Blue Diamond HP Hood.xlsm'
$strFileNam = $P.ffpath + [char]92 + $j
$fmd5 = Get-Checksum $strFileNam
$strProvidr = "Provider=Microsoft.ACE.OLEDB.12.0"
$strDataSrc = "Data Source = $strFileNam"
$strExtend  = "Extended Properties=Excel 12.0"   ## Consider usage for HDR and IMEX ( header flag & import/export mode )
$objFormConn    = New-Object System.Data.OleDb.OleDbConnection("$strProvidr;$strDataSrc;$strExtend")
$objFormConn.Open()

## $objFormConn.Close()
## Stage a command

$strQuery   = @"
Select [cYear], [cFr_Date], [cTo_Date], [cFamily]
, [cBrand], [cItem_Desc], [cCase_Pk], [cSz], [cUoM], [cUPC]
, [cList_Cost], [cCase_Freight], [cFOB_Flag]
, [cProgram_Type], [cPct_Off], [cSB_PCT], [cSB_AMT], [cREG_SRP], [cEDL_SRP]
From [EDLC Agreement Detail$]
Where [cUPC] is not NULL
And [cSz] <> -37
And [cItem_Desc] not like 'EXAMPLE -%'
"@

$sqlCommand = New-Object System.Data.OleDb.OleDbCommand($strQuery)
$sqlCommand.Connection  = $objFormConn
$sqlCommand.CommandText = $strQuery

$objAdapter = New-Object "System.Data.OleDb.OleDbDataAdapter"
$objAdapter.SelectCommand = $sqlCommand

$DataTable  = New-Object "System.Data.DataTable"
$feedback   = $objAdapter.Fill($DataTable)

$feedback.ToString().PadLeft(3) + " - $j"
# $DataTable.Rows.cREG_SRP # | Out-File -Encoding ascii -FilePath C:\tmp\upc.txt
# $DataTable.Rows[0].cSB_PCT | gm
# [decimal]($DataTable.Rows[0].cREG_SRP -replace [char]36, "")
# [Decimal]$DataTable.Rows[0].cREG_SRP.Trim([char]36)

$sqlCommand.Dispose()
$objFormConn.Close()
$objFormConn.Dispose()

# $DataTable.Rows[6].cREG_SRP -match '\d+\.?\d+'   \d*\.?\d*?

ForEach ($r in $DataTable.Rows) {
# If ( IsNullOrWhi$r.cCase_Freight) ) { "Y" } Else { "N" }
# ("X" + $r.cCase_Freight.ToString()).Length
# $r.cSz -match '\d+\.?\d+|\d+' | Out-Null
# $Sz = $matches[0]
# $dflag = [int](Test-UPCaHasCheckDigit $r.cUPC)
$in5edl = @"
Insert Into $($P.tbli) (EDL_Year, UPC, Eff_Date, End_Date
, Product_Family, Brand_Name, Item_Description
, Case_Pack, Item_Size, Item_UoM
, Case_List_Cost, Case_Freight, FOB_Flag
, Program_Type
, PCT_OFF, SB_PCT, SB_AMT, REG_SRP, EDL_SRP
, MD5_Checksum)
Values ( '$($r.cYear)', '$($r.cUPC)', '$($r.cFr_Date)', '$($r.cTo_Date)'
, '$($r.cFamily)', '$($r.cBrand -replace [Char]39, ([Char]39 + [Char]39))', '$($r.cItem_Desc -replace [Char]39, ([Char]39 + [Char]39))'
, $($r.cCase_Pk)
, $( If (("X" + $r.cSz.ToString()).Length -gt 1) { [math]::Round($r.cSz,3) } Else { 0 } )
, '$($r.cUoM)'
, $( If (("X" + $r.cList_Cost.ToString()).Length -gt 1) { [math]::Round($r.cList_Cost,2) } Else { 0 } )
, $( If (("X" + $r.cCase_Freight.ToString()).Length -gt 1) { [math]::Round($r.cCase_Freight,2) } Else { 0 } )
, '$($r.cFOB_Flag)'
, '$($r.cProgram_Type)'
, $( If (("X" + $r.cPct_Off.ToString()).Length -gt 1) { [math]::Round($r.cPct_Off,4) } Else { 0 } )
, $( If (("X" +  $r.cSB_PCT.ToString()).Length -gt 1) { [math]::Round($r.cSB_PCT,4)  } Else { 0 } )
, $( If (("X" +  $r.cSB_AMT.ToString()).Length -gt 1) { [math]::Round($r.cSB_AMT,2)  } Else { 0 } )
, $( If (("X" + $r.cREG_SRP.ToString()).Length -gt 1) { [math]::Round($r.cREG_SRP,2) } Else { 0 } )
, $( If (("X" + $r.cEDL_SRP.ToString()).Length -gt 1) { [math]::Round($r.cEDL_SRP,2) } Else { 0 } )
, '$($fmd5)'
)
"@
#$in5edl
$sqlEDLcmd = New-Object System.Data.OleDb.OleDbCommand($in5edl)
$sqlEDLcmd.Connection  = $objRepoConn
$sqlEDLcmd.CommandText = $in5edl
$feedback = $sqlEDLcmd.ExecuteNonQuery()
# $feedback
$Sz = $null
$sqlEDLcmd.Dispose()

}

}

$objRepoConn.Close()

<##############################################################

             _   _ _               _              _ 
 ___ _ _  __| | (_) |_ ___ _ __   | |___  __ _ __| |
/ -_) ' \/ _` | | |  _/ -_) '  \  | / _ \/ _` / _` |
\___|_||_\__,_| |_|\__\___|_|_|_| |_\___/\__,_\__,_|


##############################################################>

<#
$post = '\\wfm-team\team\RegionalPurchasing\National Promotions\Analytics Data\EDL_Grocery.accdb'
Copy-Item $P.strdb -Destination $post -Force
Set-ItemProperty $post -Name IsReadOnly -Value $true
#>
#### KTHXBYE ####