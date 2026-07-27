# Export-GCCPermissions.ps1
# Exports group permissions from GCC SharePoint using cookie auth.
# Set before running:
#   $env:GCC_RTFA    = "..."
#   $env:GCC_FEDAUTH = "..."

$GCCSiteUrl  = "https://cresearch3.sharepoint.com/sites/Collaborate"
$LibraryName = "CRBFS1"

$FoldersToScan = @(
    "Projects136"
)

$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$OutputDir  = Join-Path $ScriptRoot "output"
$OutCSV     = Join-Path $OutputDir "gcc_perms.csv"
$OutHTML    = Join-Path $OutputDir "gcc_perms.html"

if (-not (Test-Path $OutputDir)) { New-Item -ItemType Directory $OutputDir | Out-Null }

$rtFa    = $env:GCC_RTFA
$fedAuth = $env:GCC_FEDAUTH
if (-not $rtFa -or -not $fedAuth) { Write-Error "Set GCC_RTFA and GCC_FEDAUTH env vars first."; exit 1 }

# Build cookie session
$session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
$session.Cookies.Add((New-Object System.Net.Cookie("rtFa",    $rtFa,    "/", "sharepoint.com")))
$session.Cookies.Add((New-Object System.Net.Cookie("FedAuth", $fedAuth, "/", "cresearch3.sharepoint.com")))

$headers = @{
    "Accept"       = "application/json;odata=verbose"
    "Content-Type" = "application/json;odata=verbose"
}

# ── Get library root URL ──────────────────────────────────────
Write-Host "Getting library info..." -ForegroundColor Gray
try {
    $libCheck = Invoke-RestMethod `
        -Uri "$GCCSiteUrl/_api/web/lists/getbytitle('$LibraryName')?`$select=Title,ItemCount,RootFolder/ServerRelativeUrl&`$expand=RootFolder" `
        -Headers $headers -WebSession $session
    $libraryRootUrl = $libCheck.d.RootFolder.ServerRelativeUrl
    Write-Host "  Library: '$($libCheck.d.Title)'  items: $($libCheck.d.ItemCount)" -ForegroundColor Green
} catch {
    Write-Error "Library '$LibraryName' not found. Error: $_"; exit 1
}
# ─────────────────────────────────────────────────────────────

function Add-GroupPermissions {
    param([int]$ItemId, [string]$Path, [string]$Type,
          [System.Collections.Generic.List[PSCustomObject]]$RowsList)

    try {
        $raResult = Invoke-RestMethod `
            -Uri "$GCCSiteUrl/_api/web/lists/getbytitle('$LibraryName')/items($ItemId)/roleassignments?`$expand=Member,RoleDefinitionBindings" `
            -Headers $headers -WebSession $session

        foreach ($ra in $raResult.d.results) {
            if ($ra.Member.PrincipalType -eq 1) { continue }   # skip individual users
            $perms = ($ra.RoleDefinitionBindings.results | ForEach-Object { $_.Name }) -join "; "
            $RowsList.Add([PSCustomObject]@{
                Path          = $Path
                Type          = $Type
                Principal     = $ra.Member.Title
                PrincipalType = switch ($ra.Member.PrincipalType) {
                    2 { "DistributionList" }
                    4 { "SecurityGroup" }
                    8 { "SharePointGroup" }
                    default { "Type$($ra.Member.PrincipalType)" }
                }
                Permissions   = $perms
            })
        }
    } catch {
        Write-Warning "Role assignments failed for item $ItemId ($Path): $_"
    }
}

$rows = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($folder in $FoldersToScan) {
    $rootFolderUrl = "$libraryRootUrl/$folder"
    Write-Host "`nScanning: $rootFolderUrl" -ForegroundColor Yellow

    # Iterative breadth-first walk using the Folders API (confirmed working)
    $queue = [System.Collections.Generic.Queue[string]]::new()
    $queue.Enqueue($rootFolderUrl)

    while ($queue.Count -gt 0) {
        $currentUrl   = $queue.Dequeue()
        $encodedUrl   = $currentUrl.Replace("'", "''")

        Write-Host "  Folder: $currentUrl" -ForegroundColor Gray

        # ── Check the folder itself for unique permissions ────
        try {
            $folderResp = Invoke-RestMethod `
                -Uri "$GCCSiteUrl/_api/web/getFolderByServerRelativeUrl('$encodedUrl')?`$expand=ListItemAllFields" `
                -Headers $headers -WebSession $session

            $li = $folderResp.d.ListItemAllFields
            $hasUnique = $li.HasUniqueRoleAssignments
            Write-Host "    [diag] folder ListItem Id=$($li.Id)  HasUniqueRoleAssignments=$hasUnique" -ForegroundColor DarkGray

            if ($li -and $li.Id -and $hasUnique) {
                Write-Host "    [unique] $currentUrl" -ForegroundColor DarkYellow
                Add-GroupPermissions -ItemId $li.Id -Path $currentUrl -Type "Folder" -RowsList $rows
            }
        } catch {
            Write-Warning "Could not get folder item for $currentUrl : $_"
        }

        # ── Queue sub-folders ─────────────────────────────────
        $sfUrl = "$GCCSiteUrl/_api/web/getFolderByServerRelativeUrl('$encodedUrl')/Folders" +
                 "?`$select=Name,ServerRelativeUrl&`$expand=ListItemAllFields&`$top=500"
        $sfCount = 0
        while ($sfUrl) {
            try {
                $sfResp = Invoke-RestMethod -Uri $sfUrl -Headers $headers -WebSession $session
                foreach ($sf in $sfResp.d.results) {
                    if ($sf.Name -in @("Forms", "_t")) { continue }   # skip system folders
                    $sfCount++
                    $queue.Enqueue($sf.ServerRelativeUrl)
                }
                $sfUrl = $sfResp.d.__next
            } catch {
                Write-Warning "Could not list sub-folders of $currentUrl : $_"
                break
            }
        }
        Write-Host "    [diag] sub-folders queued: $sfCount" -ForegroundColor DarkGray

        # ── Check files in this folder ────────────────────────
        $fUrl = "$GCCSiteUrl/_api/web/getFolderByServerRelativeUrl('$encodedUrl')/Files" +
                "?`$select=ServerRelativeUrl,Name&`$expand=ListItemAllFields&`$top=500"
        $fCount = 0
        while ($fUrl) {
            try {
                $fResp = Invoke-RestMethod -Uri $fUrl -Headers $headers -WebSession $session
                foreach ($f in $fResp.d.results) {
                    $li = $f.ListItemAllFields
                    $fCount++
                    $hasUnique = $li.HasUniqueRoleAssignments
                    Write-Host "    [diag] file '$($f.Name)'  Id=$($li.Id)  HasUniqueRoleAssignments=$hasUnique" -ForegroundColor DarkGray
                    if ($li -and $li.Id -and $hasUnique) {
                        Write-Host "    [unique] $($f.ServerRelativeUrl)" -ForegroundColor DarkYellow
                        Add-GroupPermissions -ItemId $li.Id -Path $f.ServerRelativeUrl -Type "File" -RowsList $rows
                    }
                }
                $fUrl = $fResp.d.__next
            } catch {
                Write-Warning "Could not list files in $currentUrl : $_"
                break
            }
        }
        Write-Host "    [diag] files found: $fCount" -ForegroundColor DarkGray
    }
}

$rows | Export-Csv -Path $OutCSV -NoTypeInformation -Encoding UTF8
Write-Host "`nExported $($rows.Count) rows -> $OutCSV" -ForegroundColor Green

# HTML report
$tableRows = $rows | ForEach-Object {
    "<tr><td>$($_.Path)</td><td>$($_.Type)</td><td>$($_.Principal)</td><td>$($_.PrincipalType)</td><td>$($_.Permissions)</td></tr>"
}

$html = @"
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>GCC SharePoint Group Permissions</title>
  <style>
    body  { font-family: Segoe UI, Arial, sans-serif; font-size:13px; margin:24px; }
    h1    { color:#333; }
    table { border-collapse:collapse; width:100%; }
    th    { background:#005599; color:#fff; padding:8px; text-align:left; }
    td    { padding:6px 8px; border-bottom:1px solid #ddd; word-break:break-word; }
    tr:hover td { background:#f0f4ff; }
  </style>
</head>
<body>
  <h1>GCC Collaborate — Group Permissions</h1>
  <p>Folders scanned: $($FoldersToScan -join ', ') &nbsp;|&nbsp; Total rows: $($rows.Count)</p>
  <table>
    <tr><th>Path</th><th>Type</th><th>Principal</th><th>PrincipalType</th><th>Permissions</th></tr>
    $($tableRows -join "`n")
  </table>
</body>
</html>
"@

$html | Out-File -FilePath $OutHTML -Encoding UTF8
Write-Host "HTML   -> $OutHTML" -ForegroundColor Green
