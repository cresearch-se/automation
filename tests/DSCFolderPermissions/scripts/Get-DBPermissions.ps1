# Get-DBPermissions.ps1
# Connects to SQLT4COSTAGDW / DSCCaseFolder using Windows Auth,
# runs the permissions query, and exports results to CSV.
# Run this on any machine that has SQL access to SQLT4COSTAGDW.
#
# PREREQUISITES
# -------------
# 1. PowerShell 5.1 or later.
# 2. .NET SqlClient is built into PowerShell — no extra modules needed.
# 3. Windows Auth access to SQLT4COSTAGDW (domain login, no username/password required).
#    Test first: Test-NetConnection -ComputerName SQLT4COSTAGDW -Port 1433
# 4. Execution policy: powershell -ExecutionPolicy Bypass -File .\Get-DBPermissions.ps1

param(
    [Parameter(Mandatory = $false)]
    [string]$Server   = "SQLT4COSTAGDW",

    [Parameter(Mandatory = $false)]
    [string]$Database = "DSCCaseFolder",

    [Parameter(Mandatory = $false)]
    [string]$OutputFile = ""
)

# Resolve output path inside the script body so $MyInvocation is populated
if (-not $OutputFile) {
    $OutputFile = Join-Path (Split-Path $MyInvocation.MyCommand.Path -Parent) "..\output\db_permissions.csv"
    $OutputFile = [System.IO.Path]::GetFullPath($OutputFile)
}

# ---------------------------------------------------------------------------
# TODO: Replace this query with the actual query once provided
# The query should return columns that map to:
#   - Project folder name  (e.g. ProjectFolder, ProjectCode, etc.)
#   - Case folder name     (e.g. CaseFolder, CaseName, etc.)
#   - User / Group         (the identity that has access)
#   - Permission / Role    (what access they have)
# ---------------------------------------------------------------------------
$query = @"
-- REPLACE THIS WITH THE ACTUAL QUERY
SELECT TOP 10 'PLACEHOLDER' AS Note
"@

# ---------------------------------------------------------------------------
# Run query via .NET SqlClient (Windows Auth — no credentials needed)
# ---------------------------------------------------------------------------

$connStr = "Server=$Server;Database=$Database;Integrated Security=True;TrustServerCertificate=True;"

Write-Host "Connecting to : $Server / $Database" -ForegroundColor Cyan
Write-Host "Output file   : $OutputFile" -ForegroundColor Cyan

try {
    $conn = New-Object System.Data.SqlClient.SqlConnection($connStr)
    $conn.Open()
    Write-Host "Connected successfully" -ForegroundColor Green
}
catch {
    Write-Error "Could not connect to $Server/$Database : $_"
    exit 1
}

try {
    $cmd    = $conn.CreateCommand()
    $cmd.CommandText    = $query
    $cmd.CommandTimeout = 120

    $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
    $table   = New-Object System.Data.DataTable
    $adapter.Fill($table) | Out-Null

    Write-Host "Query returned $($table.Rows.Count) rows" -ForegroundColor Green
}
catch {
    Write-Error "Query failed: $_"
    $conn.Close()
    exit 1
}
finally {
    $conn.Close()
}

# ---------------------------------------------------------------------------
# Export to CSV
# ---------------------------------------------------------------------------

$outputDir = Split-Path $OutputFile -Parent
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

$table | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8

Write-Host "Done - $($table.Rows.Count) rows written to:" -ForegroundColor Green
Write-Host "  $OutputFile" -ForegroundColor Green
