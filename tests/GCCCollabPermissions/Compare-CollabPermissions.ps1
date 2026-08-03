# ============================================================
# Compare-CollabPermissions.ps1 — REST API version (no PnP)
# Legacy: Bearer token auth
# GCC:    Cookie auth (FedAuth + rtFa)
#
# Set before running:
#   $env:LEGACY_TOKEN = "eyJ..."
#   $env:GCC_RTFA     = "gPE9..."
#   $env:GCC_FEDAUTH  = "77u/..."
# ============================================================

# ── Configuration ────────────────────────────────────────────
$LegacySiteUrl = "https://cresearch1.sharepoint.com/sites/Collaborate"
$GCCSiteUrl    = "https://cresearch3.sharepoint.com/sites/Collaborate"
$LibraryName   = "CRBFS1"

$FoldersToScan = @(
    "Projects136"
)

$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$OutputDir  = Join-Path $ScriptRoot "output"
$LegacyCSV  = Join-Path $OutputDir "legacy_perms.csv"
$GCCCSV     = Join-Path $OutputDir "gcc_perms.csv"
$DiffCSV    = Join-Path $OutputDir "DIFF_REPORT.csv"
$DiffHTML   = Join-Path $OutputDir "DIFF_REPORT.html"
# ─────────────────────────────────────────────────────────────

function Build-BearerHeaders {
    param([string]$Token, [string]$Digest = $null)
    $h = @{
        "Authorization" = "Bearer $Token"
        "Accept"        = "application/json;odata=verbose"
        "Content-Type"  = "application/json;odata=verbose"
    }
    if ($Digest) { $h["X-RequestDigest"] = $Digest }
    return $h
}

function Build-CookieSession {
    param([string]$rtFa, [string]$FedAuth, [string]$Host)
    $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
    $session.Cookies.Add((New-Object System.Net.Cookie("rtFa",    $rtFa,    "/", "sharepoint.com")))
    $session.Cookies.Add((New-Object System.Net.Cookie("FedAuth", $FedAuth, "/", $Host)))
    return $session
}

function Get-FormDigest {
    param($SiteUrl, $Headers, $Session)
    $params = @{ Uri = "$SiteUrl/_api/contextinfo"; Method = "POST"; Headers = $Headers }
    if ($Session) { $params.WebSession = $Session }
    $r = Invoke-RestMethod @params -ErrorAction Stop
    return $r.d.GetContextWebInformation.FormDigestValue
}

function Invoke-SPGet {
    param($Url, $Headers, $Session)
    $params = @{ Uri = $Url; Method = "GET"; Headers = $Headers }
    if ($Session) { $params.WebSession = $Session }
    return Invoke-RestMethod @params -ErrorAction Stop
}

function Invoke-SPPost {
    param($Url, $Body, $Headers, $Session)
    $params = @{ Uri = $Url; Method = "POST"; Headers = $Headers; Body = $Body }
    if ($Session) { $params.WebSession = $Session }
    return Invoke-RestMethod @params -ErrorAction Stop
}

function Export-GroupPermissions {
    param(
        [string]$SiteUrl,
        [string]$Library,
        [string[]]$Folders,
        [string]$OutFile,
        [string]$TenantLabel,
        [hashtable]$Headers,
        $Session
    )

    Write-Host "`n=== Scanning $TenantLabel ===" -ForegroundColor Cyan

    # Get request digest for POST calls
    $digest = Get-FormDigest -SiteUrl $SiteUrl -Headers $Headers -Session $Session
    $postHeaders = $Headers.Clone()
    $postHeaders["X-RequestDigest"] = $digest

    $rows = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($folder in $Folders) {
        $folderRelUrl = "/sites/Collaborate/$Library/$folder"
        Write-Host "  Scanning: $folderRelUrl" -ForegroundColor Yellow

        # CAML query — get all items recursively in the folder
        $position = $null
        $pageCount = 0

        do {
            $positionXml = ""
            if ($position) {
                $positionXml = "<ListItemCollectionPosition PagingInfo=`"$position`" />"
            }

            $camlBody = @{
                query = @{
                    "__metadata"             = @{ "type" = "SP.CamlQuery" }
                    "FolderServerRelativeUrl" = $folderRelUrl
                    "ViewXml"                = "<View Scope='RecursiveAll'>$positionXml<ViewFields><FieldRef Name='ID'/><FieldRef Name='FileRef'/><FieldRef Name='FSObjType'/><FieldRef Name='HasUniqueRoleAssignments'/></ViewFields><RowLimit Paged='TRUE'>200</RowLimit></View>"
                }
            } | ConvertTo-Json -Depth 10

            $result = Invoke-SPPost `
                -Url "$SiteUrl/_api/web/lists/getbytitle('$Library')/GetItems" `
                -Body $camlBody `
                -Headers $postHeaders `
                -Session $Session

            $items    = $result.d.results
            $position = $result.d.__next

            Write-Host "    Page $($pageCount+1): $($items.Count) items" -ForegroundColor Gray
            $pageCount++

            foreach ($item in $items) {
                if (-not $item.HasUniqueRoleAssignments) { continue }

                $itemId   = $item.Id
                $path     = $item.FileRef
                $itemType = if ($item.FSObjType -eq 1) { "Folder" } else { "File" }

                $raResult = Invoke-SPGet `
                    -Url "$SiteUrl/_api/web/lists/getbytitle('$Library')/items($itemId)/roleassignments?`$expand=Member,RoleDefinitionBindings" `
                    -Headers $Headers `
                    -Session $Session

                foreach ($ra in $raResult.d.results) {
                    # Skip individual users (PrincipalType 1) — keep groups only (2,4,8)
                    if ($ra.Member.PrincipalType -eq 1) { continue }

                    $permLevels = ($ra.RoleDefinitionBindings.results | ForEach-Object { $_.Name }) -join "; "

                    $rows.Add([PSCustomObject]@{
                        Path          = $path
                        Type          = $itemType
                        Principal     = $ra.Member.Title
                        PrincipalType = switch ($ra.Member.PrincipalType) {
                            2  { "DistributionList" }
                            4  { "SecurityGroup" }
                            8  { "SharePointGroup" }
                            default { "Type$($ra.Member.PrincipalType)" }
                        }
                        Permissions   = $permLevels
                    })
                }
            }

        } while ($position)
    }

    $rows | Export-Csv -Path $OutFile -NoTypeInformation -Encoding UTF8
    Write-Host "  Exported $($rows.Count) group permission rows -> $OutFile" -ForegroundColor Green
}

function Compare-Permissions {
    param($LegacyFile, $GCCFile, $OutCSV, $OutHTML)

    Write-Host "`n=== Comparing permissions ===" -ForegroundColor Cyan

    $legacy = Import-Csv $LegacyFile
    $gcc    = Import-Csv $GCCFile

    $legacyMap = @{}
    foreach ($r in $legacy) { $legacyMap["$($r.Path)|$($r.Principal)"] = $r }

    $gccMap = @{}
    foreach ($r in $gcc) { $gccMap["$($r.Path)|$($r.Principal)"] = $r }

    $diff = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($key in $legacyMap.Keys) {
        $leg = $legacyMap[$key]
        if ($gccMap.ContainsKey($key)) {
            if ($leg.Permissions -ne $gccMap[$key].Permissions) {
                $diff.Add([PSCustomObject]@{
                    Path          = $leg.Path
                    Type          = $leg.Type
                    Principal     = $leg.Principal
                    PrincipalType = $leg.PrincipalType
                    Issue         = "Permission Mismatch"
                    LegacyPerms   = $leg.Permissions
                    GCCPerms      = $gccMap[$key].Permissions
                })
            }
        } else {
            $diff.Add([PSCustomObject]@{
                Path          = $leg.Path
                Type          = $leg.Type
                Principal     = $leg.Principal
                PrincipalType = $leg.PrincipalType
                Issue         = "Missing in GCC"
                LegacyPerms   = $leg.Permissions
                GCCPerms      = ""
            })
        }
    }

    foreach ($key in $gccMap.Keys) {
        if (-not $legacyMap.ContainsKey($key)) {
            $r = $gccMap[$key]
            $diff.Add([PSCustomObject]@{
                Path          = $r.Path
                Type          = $r.Type
                Principal     = $r.Principal
                PrincipalType = $r.PrincipalType
                Issue         = "Extra in GCC"
                LegacyPerms   = ""
                GCCPerms      = $r.Permissions
            })
        }
    }

    $diff | Export-Csv -Path $OutCSV -NoTypeInformation -Encoding UTF8

    $missing  = ($diff | Where-Object { $_.Issue -eq "Missing in GCC" }).Count
    $mismatch = ($diff | Where-Object { $_.Issue -eq "Permission Mismatch" }).Count
    $extra    = ($diff | Where-Object { $_.Issue -eq "Extra in GCC" }).Count

    Write-Host "`n===== DIFF SUMMARY =====" -ForegroundColor White
    Write-Host "  Missing in GCC      : $missing"  -ForegroundColor Red
    Write-Host "  Permission Mismatch : $mismatch" -ForegroundColor Yellow
    Write-Host "  Extra in GCC        : $extra"    -ForegroundColor Cyan
    Write-Host "  Total issues        : $($diff.Count)"
    Write-Host "  CSV    -> $OutCSV"

    $issueColor = @{
        "Missing in GCC"      = "#ffe0e0"
        "Permission Mismatch" = "#fff8e0"
        "Extra in GCC"        = "#e0f0ff"
    }

    $tableRows = $diff | ForEach-Object {
        $bg = $issueColor[$_.Issue]
        "<tr style='background:$bg'>
            <td>$($_.Path)</td><td>$($_.Type)</td><td>$($_.Principal)</td>
            <td>$($_.PrincipalType)</td><td><b>$($_.Issue)</b></td>
            <td>$($_.LegacyPerms)</td><td>$($_.GCCPerms)</td>
        </tr>"
    }

    $html = @"
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>GCC Collab Permissions Diff</title>
  <style>
    body  { font-family: Segoe UI, Arial, sans-serif; font-size:13px; margin:24px; }
    h1    { color:#333; }
    .summary { display:flex; gap:24px; margin-bottom:20px; }
    .card { padding:12px 20px; border-radius:6px; font-size:15px; font-weight:bold; }
    .red  { background:#ffe0e0; color:#c00; }
    .yel  { background:#fff8e0; color:#886600; }
    .blu  { background:#e0f0ff; color:#005599; }
    table { border-collapse:collapse; width:100%; }
    th    { background:#444; color:#fff; padding:8px; text-align:left; }
    td    { padding:6px 8px; border-bottom:1px solid #ddd; word-break:break-word; }
    tr:hover td { filter:brightness(0.95); }
  </style>
</head>
<body>
  <h1>GCC Collaborate Permissions Diff Report</h1>
  <p>Folders scanned: $($FoldersToScan -join ', ')</p>
  <div class="summary">
    <div class="card red">Missing in GCC: $missing</div>
    <div class="card yel">Permission Mismatch: $mismatch</div>
    <div class="card blu">Extra in GCC: $extra</div>
  </div>
  <table>
    <tr>
      <th>Path</th><th>Type</th><th>Principal</th><th>PrincipalType</th>
      <th>Issue</th><th>Legacy Permissions</th><th>GCC Permissions</th>
    </tr>
    $($tableRows -join "`n")
  </table>
</body>
</html>
"@

    $html | Out-File -FilePath $OutHTML -Encoding UTF8
    Write-Host "  HTML   -> $OutHTML" -ForegroundColor Green
}

# ── Main ─────────────────────────────────────────────────────
$LegacyToken = $env:LEGACY_TOKEN
$GCC_rtFa    = $env:GCC_RTFA
$GCC_FedAuth = $env:GCC_FEDAUTH

if (-not $LegacyToken) { Write-Error "Set `$env:LEGACY_TOKEN before running"; exit 1 }
if (-not $GCC_rtFa)    { Write-Error "Set `$env:GCC_RTFA before running"; exit 1 }
if (-not $GCC_FedAuth) { Write-Error "Set `$env:GCC_FEDAUTH before running"; exit 1 }

if (-not (Test-Path $OutputDir)) { New-Item -ItemType Directory $OutputDir | Out-Null }

# Legacy — Bearer token
$legacyHeaders = Build-BearerHeaders -Token $LegacyToken

Export-GroupPermissions `
    -SiteUrl     $LegacySiteUrl `
    -Library     $LibraryName `
    -Folders     $FoldersToScan `
    -OutFile     $LegacyCSV `
    -TenantLabel "Legacy (cresearch1)" `
    -Headers     $legacyHeaders `
    -Session     $null

# GCC — Cookie auth
$gccSession = Build-CookieSession `
    -rtFa    $GCC_rtFa `
    -FedAuth $GCC_FedAuth `
    -Host    "cresearch3.sharepoint.com"

$gccHeaders = @{
    "Accept"       = "application/json;odata=verbose"
    "Content-Type" = "application/json;odata=verbose"
}

Export-GroupPermissions `
    -SiteUrl     $GCCSiteUrl `
    -Library     $LibraryName `
    -Folders     $FoldersToScan `
    -OutFile     $GCCCSV `
    -TenantLabel "GCC (cresearch3)" `
    -Headers     $gccHeaders `
    -Session     $gccSession

Compare-Permissions `
    -LegacyFile $LegacyCSV `
    -GCCFile    $GCCCSV `
    -OutCSV     $DiffCSV `
    -OutHTML    $DiffHTML
