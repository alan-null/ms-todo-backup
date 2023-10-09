$config = Get-Content -Raw -Path "$PSScriptRoot\config.json" | ConvertFrom-Json
$client_id = $config.client_id
$client_secret = $config.client_secret
$redirect_uri = "http://localhost:8080"
$url = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?tenant=common&response_type=code&client_id=$client_id&redirect_uri=$redirect_uri&state=null&scope=Tasks.Read"

function Get-AccessToken($code) {
    $headers = @{'Content-Type' = 'application/x-www-form-urlencoded' }
    $payload = @{
        'grant_type'    = 'authorization_code'
        'client_id'     = $client_id
        'client_secret' = $client_secret
        'code'          = $code
        'redirect_uri'  = $redirect_uri
        'scope'         = "Tasks.Read"
    }
    $response = Invoke-WebRequest -Uri 'https://login.microsoftonline.com/common/oauth2/v2.0/token' `
        -Method Post `
        -Headers $headers `
        -Body $payload

    $obj = $response.Content | ConvertFrom-Json
    $obj.access_token
}

function Set-OutputStream($context, $buffer) {
    $context.Response.ContentLength64 = $buffer.Length
    $context.Response.OutputStream.Write($buffer, 0, $buffer.Length)
    $context.Response.OutputStream.Close()
}

function Export-Lists($accessToken) {
    $lists = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/me/todo/lists" `
        -Headers @{"Authorization" = "Bearer $accessToken" } `
        -Method GET

    write-host "Exporting lists to lists.json"
    $lists | ConvertTo-Json -depth 100 | Out-File "out/lists.json"

    $lists.value.id | % {
        Start-Sleep -Seconds 1
        $listID = $_
        write-host "Exporting tasks for list: $listID"
        $tasks = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/me/todo/lists/$listID/tasks" `
            -Headers @{"Authorization" = "Bearer $accessToken" } `
            -Method GET
        $tasks | ConvertTo-Json -depth 100 |  Out-File "out/$listID.json"
    }
}

$http = [System.Net.HttpListener]::new()
$http.Prefixes.Add("http://localhost:8080/")
$http.Start()

if ($http.IsListening) {
    Write-Host " HTTP Server Ready!  " -f 'black' -b 'gre'
    Write-Host "Listening $redirect_uri"
}

try {
    while ($http.IsListening) {
        $contextTask = $http.GetContextAsync()
        while (-not $contextTask.AsyncWaitHandle.WaitOne(200)) { }
        $context = $contextTask.GetAwaiter().GetResult()

        if ($context.Request.HttpMethod -eq 'GET' -and $context.Request.Url.AbsolutePath -eq '/' -and $context.Request.QueryString.AllKeys.Count -eq 0) {
            $here = "<a href='$url'>here</a>"
            [string]$html = "<h1>Microsoft To Do exporter</h1><p>Click $here to authorize and export all</p>"
            $buffer = [System.Text.Encoding]::UTF8.GetBytes($html)
            Set-OutputStream $context $buffer
        }

        if ($context.Request.HttpMethod -eq 'GET' -and $context.Request.Url.AbsolutePath.TrimStart('/') -eq '' -and $context.Request.QueryString.AllKeys.Contains("code")) {
            $code = $context.Request.QueryString['code']

            $accessToken = Get-AccessToken $code

            $buffer = [System.Text.Encoding]::UTF8.GetBytes("Exporting in the background...")
            Set-OutputStream $context $buffer

            Export-Lists $accessToken

            $http.Stop()
        }

        if ($context.Request.HttpMethod -eq 'GET' -and $context.Request.QueryString.AllKeys.Contains("error")) {
            [string]$html = $context.Request.QueryString['error_description']
            $buffer = [System.Text.Encoding]::UTF8.GetBytes($html)
            Set-OutputStream $context $buffer
        }
    }
}
finally {
    $http.Stop()
}