# Compare-Permissions.ps1
# Compares folder permissions from Windows (Get-FolderPermissions.ps1 output)
# against permissions from the database (Get-DBPermissions.ps1 output).
#
# Produces:
#   - Console summary of mismatches
#   - output\comparison_report.csv  with each difference tagged by type
#
# PREREQUISITES
# -------------
# 1. PowerShell 5.1 or later.
# 2. Both input CSVs must exist (run Get-FolderPermissions.ps1 and Get-DBPermissions.ps1 first).
# 3. Execution policy: powershell -ExecutionPolicy Bypass -File .\Compare-Permissions.ps1

param(
    [Parameter(Mandatory = $false)]
    [string]$FolderPermissionsFile = "",   # output of Get-FolderPermissions.ps1

    [Parameter(Mandatory = $false)]
    [string]$DBPermissionsFile = "",       # output of Get-DBPermissions.ps1

    [Parameter(Mandatory = $false)]
    [string]$OutputFile = ""
)

# Resolve default paths inside script body
$scriptDir = Split-Path $MyInvocation.MyCommand.Path -Parent
$outputDir = [System.IO.Path]::GetFullPath((Join-Path $scriptDir "..\output"))

if (-not $FolderPermissionsFile) {
    $FolderPermissionsFile = Join-Path $outputDir "folder_permissions.csv"
}
if (-not $DBPermissionsFile) {
    $DBPermissionsFile = Join-Path $outputDir "db_permissions.csv"
}
if (-not $OutputFile) {
    $OutputFile = Join-Path $outputDir "comparison_report.csv"
}

# ---------------------------------------------------------------------------
# TODO: Configure these column mappings once the DB query columns are known.
# These tell the script which column in each CSV represents the same concept.
#
# Example — if the folder CSV uses "FolderName" and the DB CSV uses "CaseName":
#   $MAP_CASE_FOLDER   = @{ Folder = "FolderName";   DB = "CaseName" }
#   $MAP_PROJECT       = @{ Folder = "ParentProject"; DB = "ProjectFolder" }
#   $MAP_IDENTITY      = @{ Folder = "Identity";      DB = "UserOrGroup" }
#   $MAP_RIGHTS        = @{ Folder = "FileSystemRights"; DB = "PermissionLevel" }
#
# Update the right-hand side values below once the DB query is provided.
# ---------------------------------------------------------------------------
$MAP_CASE_FOLDER = @{ Folder = "FolderName";       DB = "TODO_DB_CASE_COLUMN" }
$MAP_PROJECT     = @{ Folder = "ParentProject";     DB = "TODO_DB_PROJECT_COLUMN" }
$MAP_IDENTITY    = @{ Folder = "Identity";          DB = "TODO_DB_IDENTITY_COLUMN" }
$MAP_RIGHTS      = @{ Folder = "FileSystemRights";  DB = "TODO_DB_RIGHTS_COLUMN" }

# ---------------------------------------------------------------------------
# Load CSVs
# ---------------------------------------------------------------------------

foreach ($f in @($FolderPermissionsFile, $DBPermissionsFile)) {
    if (-not (Test-Path $f)) {
        Write-Error "File not found: $f"
        exit 1
    }
}

Write-Host "Loading folder permissions : $FolderPermissionsFile" -ForegroundColor Cyan
Write-Host "Loading DB permissions     : $DBPermissionsFile" -ForegroundColor Cyan

$folderData = Import-Csv $FolderPermissionsFile
$dbData     = Import-Csv $DBPermissionsFile

# Case folders only — project-level comparison is separate
$folderCases = $folderData | Where-Object { $_.FolderType -eq "Case" }

Write-Host "Folder CSV  : $($folderCases.Count) case-folder permission rows" -ForegroundColor Green
Write-Host "DB CSV      : $($dbData.Count) rows" -ForegroundColor Green

# ---------------------------------------------------------------------------
# Build lookup key: ProjectFolder|CaseFolder|Identity  ->  Rights
# ---------------------------------------------------------------------------

function Make-Key($project, $caseName, $identity) {
    return "$($project.Trim().ToUpper())|$($caseName.Trim().ToUpper())|$($identity.Trim().ToUpper())"
}

$folderIndex = @{}
foreach ($row in $folderCases) {
    $key = Make-Key $row.($MAP_PROJECT.Folder) $row.($MAP_CASE_FOLDER.Folder) $row.($MAP_IDENTITY.Folder)
    $folderIndex[$key] = $row.($MAP_RIGHTS.Folder)
}

$dbIndex = @{}
foreach ($row in $dbData) {
    $key = Make-Key $row.($MAP_PROJECT.DB) $row.($MAP_CASE_FOLDER.DB) $row.($MAP_IDENTITY.DB)
    $dbIndex[$key] = $row.($MAP_RIGHTS.DB)
}

# ---------------------------------------------------------------------------
# Compare
# ---------------------------------------------------------------------------

$results = [System.Collections.Generic.List[PSCustomObject]]::new()

# Entries in folders but missing or different in DB
foreach ($key in $folderIndex.Keys) {
    $parts  = $key -split '\|'
    $project = $parts[0]; $caseName = $parts[1]; $identity = $parts[2]

    if (-not $dbIndex.ContainsKey($key)) {
        $results.Add([PSCustomObject]@{
            DiffType      = "MISSING_IN_DB"
            Project       = $project
            CaseFolder    = $caseName
            Identity      = $identity
            FolderRights  = $folderIndex[$key]
            DBRights      = ""
        })
    }
    elseif ($folderIndex[$key] -ne $dbIndex[$key]) {
        $results.Add([PSCustomObject]@{
            DiffType      = "RIGHTS_MISMATCH"
            Project       = $project
            CaseFolder    = $caseName
            Identity      = $identity
            FolderRights  = $folderIndex[$key]
            DBRights      = $dbIndex[$key]
        })
    }
}

# Entries in DB but missing in folders
foreach ($key in $dbIndex.Keys) {
    if (-not $folderIndex.ContainsKey($key)) {
        $parts  = $key -split '\|'
        $results.Add([PSCustomObject]@{
            DiffType      = "MISSING_IN_FOLDER"
            Project       = $parts[0]
            CaseFolder    = $parts[1]
            Identity      = $parts[2]
            FolderRights  = ""
            DBRights      = $dbIndex[$key]
        })
    }
}

# ---------------------------------------------------------------------------
# Summary + Export
# ---------------------------------------------------------------------------

$missingInDB     = @($results | Where-Object { $_.DiffType -eq "MISSING_IN_DB" }).Count
$missingInFolder = @($results | Where-Object { $_.DiffType -eq "MISSING_IN_FOLDER" }).Count
$mismatch        = @($results | Where-Object { $_.DiffType -eq "RIGHTS_MISMATCH" }).Count

Write-Host ""
Write-Host "=== Comparison Summary ===" -ForegroundColor Yellow
Write-Host "  MISSING_IN_DB     : $missingInDB"     -ForegroundColor Red
Write-Host "  MISSING_IN_FOLDER : $missingInFolder" -ForegroundColor Red
Write-Host "  RIGHTS_MISMATCH   : $mismatch"         -ForegroundColor Red
Write-Host "  TOTAL DIFFS       : $($results.Count)"

if ($results.Count -eq 0) {
    Write-Host "  All permissions match!" -ForegroundColor Green
}

$results | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8
Write-Host ""
Write-Host "Report written to: $OutputFile" -ForegroundColor Green
