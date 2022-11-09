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
$Found | Select-Object -Property id, server, created_at

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


[System.Collections.Generic.List[PSObject]] $TopAccounts = @()
$MonthlyAccount = Group-ByDate -Data $MonthData -DatePattern 'yyyy-MM' -GroupObjects ('Date', 'acct_url')
$TopAccounts.Add("# Top Users for $((Get-Culture).DateTimeFormat.GetMonthName($EndDate.Month)) $($EndDate.Year)")
$TopAccounts.Add('| User | Display Name | Server | Post |')
$TopAccounts.Add('| -- | -- | -- | -- |')
$MonthlyAccount | ForEach-Object{ 
    $TopAccounts.Add("| [$($_.username)]($($_.acct_url)) | $($_.DisplayName) | $($_.server) | $($_.Posts) |")
}
$TopAccounts | Out-File (Join-Path $Path 'Reports\TopAccounts.md') -Encoding utf8