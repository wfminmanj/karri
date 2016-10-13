# ~cr~repo~\cr_repo_.xml or _dev_repo_\_dev_repo_.xml
$P = @{
'message_repo' = '\\wfm-team\team\RegionalPurchasing\National Promotions\_dev_repo_\_dev_repo_.xml'
'opening' = @"
<html><head><meta http-equiv=Content-Type content="text/html; charset=iso-8859-1"></head><body lang=EN-US>
<div style="font-family:helvetica"><h1 style='color:#6600CC'>
Grocery Category Review Files
</h1><h3>
{subject}
</h3><a href="{href}">Original Message</a>
  <h3>Attachments:</h3><ul>
"@
'closing' = @"
</ul></div></body></html>
"@
}
If (Test-Path $P.message_repo) {
$X = New-Object System.Xml.XmlDocument
$X.Load($P.message_repo) }
Else { Write-Host -Back Black -Fore Red "ERROR attempting to set repo as $($P.message_repo)" }

$routes = @{
    "/favicon.ico" = {'ignored'}
    "/ola" = { return '<html><body>Hello world!</body></html>' }
    "/stop" = { break }
    "/linklkup" = {
    $msgkey = $($requestUrl.Query -Split [char]61)[-1]
    $nodemsg = $X.SelectSingleNode("/repo/messages/msg[@msgkey='$([System.Security.SecurityElement]::Escape($msgkey))']")

    $htmlout = $P.opening -Replace '{subject}', $nodemsg.subject
    $htmlout = $htmlout -Replace '{href}', $nodemsg.href
    ForEach ( $f in $nodemsg.Attachments.ChildNodes ) {
<# 2017-AUG-19 Jeff I: replace escape logic with EscapeDataString logic
        $htmlout = $htmlout + "<li><a href=""$($f.href)"" >$([System.Security.SecurityElement]::Escape($f.name))</a></li>" }
#>
        $htmlout = $htmlout + "<li><a href=""$($f.href)"" >$([System.URI]::EscapeDataString($f.name))</a></li>" }
    $htmlout = $htmlout + $P.closing
    Return $htmlout
     }
}

$url = 'http://localhost:8090/'
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add($url)
$listener.Start()

Write-Host "Listening at $url..."

While ($listener.IsListening)
{
    $context = $listener.GetContext()
    $requestUrl = $context.Request.Url
    $requestQry = $context.Request.QueryString
    $response = $context.Response

    Write-Host ''
    Write-Host "> $requestUrl"

    $localPath = $requestUrl.LocalPath
    $route = $routes.Get_Item($requestUrl.LocalPath)

    if ($route -eq $null)
    {
        $response.StatusCode = 404
    }
    else
    {
        $content = & $route
        $buffer = [System.Text.Encoding]::UTF8.GetBytes($content)
        $response.ContentLength64 = $buffer.Length
        $response.OutputStream.Write($buffer, 0, $buffer.Length)
    }

    $response.Close()

    $responseStatus = $response.StatusCode
    Write-Host "< $responseStatus"
}

#### KTHXBYE ####
