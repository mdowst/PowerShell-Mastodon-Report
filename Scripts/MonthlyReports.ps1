#$LastRunCsv = 'C:\Users\mdowst\OneDrive - Catapult Systems\Scripts\Projects\PowerShell-Mastodon-Report\Data\LastRun.csv'
$Path = Split-Path $PSScriptRoot
$EndDate = (Get-Date -Day 1).Date
$servers = Get-Content -LiteralPath (Join-Path $Path 'Lists\PoshServers.txt')
$LastRunCsv = Join-Path $Path "Data\$($EndDate.ToString('yyyy-MM')).csv"
if(Test-Path $LastRunCsv){
    $PastData = Import-Csv $LastRunCsv
}
else{
    $PastData = $null
}
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

Function Group-ByDate{
    param(
        $Data,
        $DatePattern='yyyy-MM',
        $GroupObjects = @('acct_url', 'Date')
    )
    $Grouped = $Data | Select-Object -Property *, @{l='Date';e={$_.created_at.ToString($DatePattern)}} | Group-Object $GroupObjects
        
    $Summarized = $Grouped | Select-Object -Property Count, @{l='interactions';e={($_.Group.interactions | Measure-Object -Sum).Sum}}, 
        @{l='ActiveUsers';e={($_.Group.username | Select-Object -Unique | Measure-Object).Count}}, 
        @{l='Posts';e={($_.Group.id | Measure-Object).Count}}, @{l='Group';e={$_.Group[0]}} |
        Select-Object -Property Count, Interactions, ActiveUsers, Posts -ExpandProperty Group

    $Summarized | Select-Object -Property *, @{l='Score';e={$_.interactions + $_.users + $_.posts}} | Sort-Object Score -Descending
}

Function Expand-MDLink($string){
    [pscustomobject]@{
        Text = [Regex]::Match($string, '(?<=\[)(.*?)(?=])').Value
        Link = [Regex]::Match($string, '(?<=\()(.*?)(?=\))').Value
    }
}

[System.Collections.Generic.List[PSObject]] $Found = @()
$wp = 0
foreach ($srv in $servers) {
    Write-Progress -Activity "Found: $($Found.Count)" -Status "$wp of $($servers.count)" -PercentComplete $(($wp / $($servers.count)) * 100) -id 1; $wp++
    $query = @('local=true')

    $ServerData = $PastData | Where-Object{ $_.Server -eq $srv } 
    $Last = $ServerData | Sort-Object created_at | Select-Object -Last 1
    if($Last){
        $query += "since_id=$($last.id)"
        $End = $last.created_at
    }
    else{
        $End = $EndDate
    }

    try {
        do {
            $uri = "https://$($srv)/api/v1/timelines/tag/powershell?$($query -join('&'))"
            $uri
            $mast = Invoke-TimedRestMethod $uri -ErrorAction Stop
            if ($mast) {
                $mast | Select-Object *, @{l = 'searched'; e = { $srv } }, @{l = 'server'; e = { $_.uri.Split('/')[2] } } | 
                    Where-Object { ($_.tags.name -contains 'powershell' -or $_.content -match 'powershell') -and $_.id -notin $ServerData.id -and 
                    $_.created_at -ge $EndDate } | 
                    ForEach-Object { $Found.Add($_) }
                
                $lastEntry = $mast | Sort-Object created_at -Descending | Select-Object -Last 1
                if ($query | Where-Object{$_ -eq "max_id=$($lastEntry.id)"}) {
                    $lastEntry = $null
                }
                elseif($query | Where-Object{$_ -match '^max_id'}){
                    $query | Where-Object{$_ -match '^max_id'} | ForEach-Object{ 
                        $query[$query.IndexOf($_)] = "max_id=$($lastEntry.id)" 
                    }
                }
                else{
                    $query += "max_id=$($lastEntry.id)"
                }
                
            }
            else {
                $lastEntry = $null
            }
        }while ($lastEntry.created_at -gt $End)
    }
    catch {
        "$srv - $uri - $($_.Exception.Message.Split("`n")[0])"
    }
}
Write-Progress -Activity "Done" -Id 1 -Completed

$Found.Count

$Found | Select-Object -Property id, server, @{l='username';e={$_.account.username}}, @{l='DisplayName';e={$_.account.display_name}}, @{l='acct_url';e={$_.account.url}}, 
    replies_count, reblogs_count, favourites_count, url, uri, created_at | Export-Csv $LastRunCsv -Append


$MonthData = Import-Csv $LastRunCsv
$MonthData = $MonthData | Group-Object id | ForEach-Object{
    $_.Group | Select-Object -First 1
}


[System.Collections.Generic.List[PSObject]] $TopServers = @()
$MonthlyServer = Group-ByDate -Data $MonthData -DatePattern 'yyyy-MM' -GroupObjects ('Date', 'server')
$TopServers.Add("# Top Servers for $((Get-Culture).DateTimeFormat.GetMonthName($EndDate.Month)) $($EndDate.Year)")
$TopServers.Add('| Server | Posts | Active Users |')
$TopServers.Add('| -- | -- | -- |')
$MonthlyServer | ForEach-Object{
    $TopServers.Add("| [$($_.Server)](https://$($_.Server)/tags/PowerShell) | $($_.Posts) | $($_.ActiveUsers) |")
}
$TopServers | Out-File (Join-Path $Path 'Reports\TopServers.md') -Encoding utf8
$TopServers | Out-File (Join-Path $Path "Reports\Historical\$($EndDate.ToString('yyyy-MM')).TopServers.md") -Encoding utf8


[System.Collections.Generic.List[PSObject]] $TopAccounts = @()
$MonthlyAccount = Group-ByDate -Data $MonthData -DatePattern 'yyyy-MM' -GroupObjects ('Date', 'acct_url')
$TopAccounts.Add("# Top Users for $((Get-Culture).DateTimeFormat.GetMonthName($EndDate.Month)) $($EndDate.Year)")
$TopAccounts.Add('| User | Display Name | Server | Post |')
$TopAccounts.Add('| -- | -- | -- | -- |')
$MonthlyAccount | ForEach-Object{ 
    $TopAccounts.Add("| [$($_.username)]($($_.acct_url)) | $($_.DisplayName) | $($_.server) | $($_.Posts) |")
}
$TopAccounts | Out-File (Join-Path $Path 'Reports\TopAccounts.md') -Encoding utf8
$TopAccounts | Out-File (Join-Path $Path "Reports\Historical\$($EndDate.ToString('yyyy-MM')).TopAccounts.md") -Encoding utf8

# Create recent posts report
[System.Collections.Generic.List[PSObject]] $RecentPosts = @()
$Found | Where-Object{ $_.created_at -gt (Get-Date).AddDays(-1) } | Select-Object @{l='account';e={$_.account.display_name}}, @{l='accountUrl';e={$_.account.url}}, 
    created_at, url, content | ForEach-Object{ $RecentPosts.Add($_) }

$RecentPostsMDPath = Join-Path $Path 'Reports\RecentPosts.md'
if(Test-Path $RecentPostsMDPath){
    $currentRecentPosts = Get-Content $RecentPostsMDPath
    for($i = $currentRecentPosts.IndexOf('| -- | -- | -- |')+1; $i -lt $currentRecentPosts.Count; $i++){
        $split = ($currentRecentPosts[$i] -replace('^\|','') -replace('\|$','')).Split('|')
        if($split.Count -ge 3){
            $account = Expand-MDLink $split[0]
            $link = Expand-MDLink $split[1]
            $content = 2..$($split.Count-1) | ForEach-Object{
                $split[$_]
            }
            $account | Select-Object @{l='account';e={$_.Text}}, @{l='accountUrl';e={$_.Link}}, @{l='created_at';e={Get-Date $link.Text}}, @{l='url';e={$link.Link}}, 
                @{l='content';e={$content -join('|')}} | ForEach-Object{ $RecentPosts.Add($_) }
        }
    }
}
$FilteredRecentPosts = $RecentPosts | Group-Object url | ForEach-Object{
    $_.Group | Select-Object -First 1
}

[System.Collections.Generic.List[PSObject]] $RecentPostMD = @()
$RecentPostMD.Add("# Recent PowerShell Topics")
$RecentPostMD.Add('| User | Date/Link | Content |')
$RecentPostMD.Add('| -- | -- | -- |')
$FilteredRecentPosts | Sort-Object created_at -Descending | ForEach-Object{
    $RecentPostMD.Add("| [$($_.account)]($($_.accounturl)) | [$($_.created_at)]($($_.url)) | $($_.content) |")
}
$RecentPostMD | Out-File $RecentPostsMDPath -Encoding utf8