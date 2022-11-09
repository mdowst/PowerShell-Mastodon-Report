$serverList = 'C:\Users\mdowst\OneDrive - Catapult Systems\Scripts\Projects\PowerShell-Mastodon-Report\Lists\PoshServers.txt'

Function Invoke-TimedRestMethod {
    Param(
        [Parameter(Mandatory = $false)]
        [string]$uri = $null
    )


    $job = Start-Job -ScriptBlock { param($uri)
        $request = Invoke-Webrequest -URI $uri -TimeoutSec 10
        $request.Content } -ArgumentList $uri
        
    $timer = [system.diagnostics.stopwatch]::StartNew()
    while ($job.State -ne 'Completed') {
        if ($timer.Elapsed.TotalSeconds -gt 10) {
            $job | Stop-Job
            throw "timeout"
        }
    }
        
    $content = $job | Receive-Job
    try{
        $content | ConvertFrom-Json -ErrorAction Stop
    }
    catch{
        throw "JSON convert failed: $content"
    }
}

$servers = Get-Content -LiteralPath $serverList
$EndDate = (Get-Date '2022-11-08T00:00:00.000Z').ToUniversalTime()
[System.Collections.Generic.List[PSObject]] $Found = @()
$wp = 0
foreach ($srv in $servers) {
    Write-Progress -Activity "Found: $($Found.Count)" -Status "$wp of $($servers.count)" -PercentComplete $(($wp / $($servers.count)) * 100) -id 1; $wp++
    $uri = "https://$($srv)/api/v1/timelines/tag/powershell"
    try {
        do {
            $mast = Invoke-TimedRestMethod $uri -ErrorAction Stop
            if ($mast) {
                $mast | Select-Object *, @{l = 'searched'; e = { $srv } }, @{l = 'server'; e = { $_.uri.Split('/')[2] } } | 
                    Where-Object { ($_.tags.name -contains 'powershell' -or $_.content -match 'powershell') -and $_.created_at -gt $EndDate } | 
                    ForEach-Object { $Found.Add($_) }
                $lastEntry = $mast | Sort-Object created_at -Descending | Select-Object -ExpandProperty created_at -Last 1
                $uri = "https://$($server)/api/v1/timelines/public?max_id=$($lastEntry.id)"
            }
            else {
                $lastEntry = $null
            }
        }while ($lastEntry.created_at -gt $EndDate)
    }
    catch {
        "$srv - $uri - $($_.Exception.Message.Split("`n")[0])"
    }
}
Write-Progress -Activity "Done" -Id 1 -Completed

$Found.Count
$Grouped = $Found | Select-Object *, @{l = 'rank'; e = { $_.uri -match $_.searched } } | Group-Object uri | ForEach-Object {
    $_.Group | Sort-Object rank -Descending | Select-Object -First 1
}
$Grouped.Count

$Cleaned = $Grouped | Group-Object url | ForEach-Object {
    $_.Group | Sort-Object rank -Descending | Select-Object -First 1
}
$Cleaned.Count


$DailyGroups = $Cleaned | Select-Object server, @{l='username';e={$_.account.username}}, @{l='display_name';e={$_.account.display_name}}, 
    @{l='interactions';e={$_.replies_count + $_.reblogs_count + $_.favourites_count}},
    @{l='acct_url';e={$_.account.url}}, @{l='Date';e={$_.created_at.Date}}, id | Group-Object acct_url, Date 
    
$Daily = $DailyGroups | Select-Object Count, @{l='server';e={$_.Group[0].server}}, @{l='username';e={$_.Group[0].username}}, @{l='display_name';e={$_.Group[0].display_name}}, 
    @{l='acct_url';e={$_.Group[0].acct_url}}, @{l='Date';e={$_.Group[0].Date}}, @{l='interactions';e={($_.Group.interactions | Measure-Object -Sum).Sum}},
    @{l='posts';e={($_.Group.id | Measure-Object).Count}}
    
$Daily | FT Count, Server, username, interactions, posts, date, acct_url


break
    $Cleaned[0].account
$Cleaned | Where-Object { $_.server -notin $servers } | Select-Object -ExpandProperty server -Unique | Out-File -FilePath $serverList -Encoding utf8 -Append
$servers -contains 'mastodon.ie'

$srv = 'mstdn.binfalse.de'
$mast