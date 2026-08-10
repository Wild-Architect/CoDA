param(
    [ValidateSet('Initial', 'Update', 'SetUrl', 'Verify')]
    [string]$Mode = 'Update',
    [string]$Version,
    [string]$DatabaseDate,
    [string]$Url
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$projectRoot = Split-Path -Parent $root
$catalogPath = Join-Path $projectRoot 'data\dbn_catalog.xlsx'
$buildingCodesPath = Join-Path $projectRoot 'building_codes'
$manifestPath = Join-Path $root 'manifest.json'
$releases = Join-Path $root 'releases'
$script:sourceByRelativePath = @{}

function Write-JsonFile([string]$Path, $Value) {
    $json = $Value | ConvertTo-Json -Depth 12
    [System.IO.File]::WriteAllText($Path, $json, (New-Object System.Text.UTF8Encoding($false)))
}

function Get-Inventory {
    $files = @(
        [pscustomobject]@{ Source = $catalogPath; Relative = 'dbn_catalog.xlsx' }
        Get-ChildItem -LiteralPath $buildingCodesPath -File -Recurse |
            Where-Object { $_.Name -notlike '*.inspect.ndjson' } |
            Sort-Object FullName | ForEach-Object {
                [pscustomobject]@{
                    Source = $_.FullName
                    Relative = 'building_codes/' + $_.FullName.Substring($buildingCodesPath.Length + 1).Replace('\', '/')
                }
            }
    )
    return @($files | ForEach-Object {
        $script:sourceByRelativePath[$_.Relative] = $_.Source
        $item = Get-Item -LiteralPath $_.Source
        [ordered]@{
            path = $_.Relative
            size = $item.Length
            sha256 = (Get-FileHash -LiteralPath $_.Source -Algorithm SHA256).Hash.ToLowerInvariant()
        }
    })
}

function New-Stage {
    $stage = Join-Path $releases ('.stage-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $stage -Force | Out-Null
    return $stage
}

function Remove-Stage([string]$Stage) {
    $resolvedRoot = [System.IO.Path]::GetFullPath($releases) + [System.IO.Path]::DirectorySeparatorChar
    $resolvedStage = [System.IO.Path]::GetFullPath($Stage)
    if (-not $resolvedStage.StartsWith($resolvedRoot) -or -not (Split-Path -Leaf $resolvedStage).StartsWith('.stage-')) {
        throw 'Unsafe temporary directory path.'
    }
    Remove-Item -LiteralPath $resolvedStage -Recurse -Force
}

function Copy-DatabaseFile([string]$RelativePath, [string]$Stage) {
    $source = $script:sourceByRelativePath[$RelativePath]
    if (-not $source) { throw "Source file was not found for '$RelativePath'." }
    $destination = Join-Path $Stage ($RelativePath.Replace('/', '\'))
    New-Item -ItemType Directory -Path (Split-Path -Parent $destination) -Force | Out-Null
    Copy-Item -LiteralPath $source -Destination $destination
}

if (-not (Test-Path -LiteralPath $catalogPath -PathType Leaf)) { throw "Catalog was not found: $catalogPath" }
if (-not (Test-Path -LiteralPath $buildingCodesPath -PathType Container)) { throw "Building codes directory was not found: $buildingCodesPath" }

if ($Mode -eq 'SetUrl') {
    if (-not $Version) { $Version = Read-Host 'Package version' }
    if (-not $Url) { $Url = Read-Host 'Public GitHub Release download URL' }
    if (-not (Test-Path -LiteralPath $manifestPath)) { throw 'manifest.json was not found.' }
    $manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $matched = $false
    if ($manifest.fullPackage.version -eq $Version) { $manifest.fullPackage.url = $Url; $matched = $true }
    foreach ($update in $manifest.updates) {
        if ($update.to -eq $Version) { $update.url = $Url; $matched = $true }
    }
    if (-not $matched) { throw "Package version '$Version' was not found in manifest.json." }
    Write-JsonFile $manifestPath $manifest
    Write-Host 'Link saved to manifest.json.'
    exit 0
}

$inventory = Get-Inventory
if (-not ($inventory | Where-Object { $_.path -eq 'dbn_catalog.xlsx' })) { throw 'data/dbn_catalog.xlsx was not found.' }
if (-not ($inventory | Where-Object { $_.path -like 'building_codes/*' })) { throw 'No building code files were found.' }

if ($Mode -eq 'Verify') {
    Write-Host "Database source is valid: $($inventory.Count - 1) document files and one catalog."
    Write-Host "Catalog: $catalogPath"
    Write-Host "Documents: $buildingCodesPath"
    exit 0
}

if (-not $Version) { $Version = Read-Host 'New database version (example: 2026.08.09)' }
if (-not $DatabaseDate) { $DatabaseDate = Read-Host 'Database date (example: 2026-08-09)' }
if ($Version -notmatch '^[0-9A-Za-z._-]+$') { throw 'Version contains unsupported characters.' }
if ($DatabaseDate -notmatch '^\d{4}-\d{2}-\d{2}$') { throw 'Database date must use YYYY-MM-DD.' }

New-Item -ItemType Directory -Path $releases -Force | Out-Null

if ($Mode -eq 'Initial') {
    $stage = New-Stage
    try {
        foreach ($file in $inventory) { Copy-DatabaseFile $file.path $stage }
        Write-JsonFile (Join-Path $stage 'update.json') ([ordered]@{ type='full'; version=$Version; databaseDate=$DatabaseDate; delete=@() })
        $packageName = "CoDA-full-$Version.zip"
        $packagePath = Join-Path $releases $packageName
        if (Test-Path -LiteralPath $packagePath) { throw "Package already exists: $packageName" }
        Compress-Archive -Path (Join-Path $stage '*') -DestinationPath $packagePath -CompressionLevel Optimal
        $manifest = [ordered]@{
            schemaVersion = 1
            latestVersion = $Version
            databaseDate = $DatabaseDate
            generatedAt = (Get-Date).ToUniversalTime().ToString('o')
            fullPackage = [ordered]@{
                version = $Version; file = $packageName; url = ''
                size = (Get-Item -LiteralPath $packagePath).Length
                sha256 = (Get-FileHash -LiteralPath $packagePath -Algorithm SHA256).Hash.ToLowerInvariant()
            }
            updates = @()
            files = $inventory
        }
        Write-JsonFile $manifestPath $manifest
        Write-Host "Initial package created: releases\$packageName"
    } finally { if ($stage -and (Test-Path -LiteralPath $stage)) { Remove-Stage $stage } }
    exit 0
}

if (-not (Test-Path -LiteralPath $manifestPath)) { throw 'Run Initial mode first.' }
$previous = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
if ($previous.latestVersion -eq $Version) { throw 'The new version must differ from the current version.' }
$previousByPath = @{}; foreach ($file in $previous.files) { $previousByPath[$file.path] = $file.sha256 }
$currentByPath = @{}; foreach ($file in $inventory) { $currentByPath[$file.path] = $file.sha256 }
$changed = @($inventory | Where-Object { -not $previousByPath.ContainsKey($_.path) -or $previousByPath[$_.path] -ne $_.sha256 })
$deleted = @($previous.files | Where-Object { -not $currentByPath.ContainsKey($_.path) } | ForEach-Object { $_.path })
if ($changed.Count -eq 0 -and $deleted.Count -eq 0) { throw 'No database changes were found.' }

$stage = New-Stage
try {
    foreach ($file in $changed) { Copy-DatabaseFile $file.path $stage }
    Write-JsonFile (Join-Path $stage 'update.json') ([ordered]@{
        type='incremental'; from=$previous.latestVersion; to=$Version; databaseDate=$DatabaseDate; delete=$deleted
    })
    $packageName = "CoDA-update-$Version.zip"
    $packagePath = Join-Path $releases $packageName
    if (Test-Path -LiteralPath $packagePath) { throw "Package already exists: $packageName" }
    Compress-Archive -Path (Join-Path $stage '*') -DestinationPath $packagePath -CompressionLevel Optimal
    $updates = @($previous.updates) + @([ordered]@{
        from = $previous.latestVersion; to = $Version; file = $packageName; url = ''
        size = (Get-Item -LiteralPath $packagePath).Length
        sha256 = (Get-FileHash -LiteralPath $packagePath -Algorithm SHA256).Hash.ToLowerInvariant()
        changedFiles = $changed.Count; deletedFiles = $deleted.Count
    })
    $manifest = [ordered]@{
        schemaVersion = 1; latestVersion = $Version; databaseDate = $DatabaseDate
        generatedAt = (Get-Date).ToUniversalTime().ToString('o')
        fullPackage = $previous.fullPackage; updates = $updates; files = $inventory
    }
    Write-JsonFile $manifestPath $manifest
    Write-Host "Update package created: releases\$packageName"
    Write-Host "Changed files: $($changed.Count); deleted files: $($deleted.Count)"
} finally { if ($stage -and (Test-Path -LiteralPath $stage)) { Remove-Stage $stage } }
