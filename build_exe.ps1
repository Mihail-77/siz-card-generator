$ErrorActionPreference = "Stop"

$projectRoot = (Resolve-Path -LiteralPath $PSScriptRoot).Path
Set-Location -LiteralPath $projectRoot

function Assert-ProjectChildPath {
    param([Parameter(Mandatory = $true)][string]$Path)

    $fullPath = [System.IO.Path]::GetFullPath($Path)
    $rootPrefix = $projectRoot.TrimEnd("\") + "\"
    if (-not $fullPath.StartsWith($rootPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Unsafe build path: $fullPath"
    }
    return $fullPath
}

function Remove-BuildDirectory {
    param([Parameter(Mandatory = $true)][string]$RelativePath)

    $target = Assert-ProjectChildPath (Join-Path $projectRoot $RelativePath)
    if (Test-Path -LiteralPath $target) {
        Remove-Item -LiteralPath $target -Recurse -Force
    }
}

Remove-BuildDirectory "build"
Remove-BuildDirectory "dist"
Remove-BuildDirectory "release\SIZCardGenerator"

& python -m PyInstaller --clean --noconfirm "SIZCardGenerator.spec"
if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller failed with exit code $LASTEXITCODE."
}

$distApp = Assert-ProjectChildPath (Join-Path $projectRoot "dist\SIZCardGenerator")
$releaseRoot = Assert-ProjectChildPath (Join-Path $projectRoot "release")
$releaseApp = Assert-ProjectChildPath (Join-Path $releaseRoot "SIZCardGenerator")
if (-not (Test-Path -LiteralPath (Join-Path $distApp "SIZCardGenerator.exe") -PathType Leaf)) {
    throw "PyInstaller did not create SIZCardGenerator.exe."
}

New-Item -ItemType Directory -Path $releaseRoot -Force | Out-Null
Copy-Item -LiteralPath $distApp -Destination $releaseRoot -Recurse

foreach ($relativeDirectory in @("data", "templates", "output", "output\backups", "requests")) {
    New-Item -ItemType Directory -Path (Join-Path $releaseApp $relativeDirectory) -Force | Out-Null
}

Copy-Item -LiteralPath (Join-Path $projectRoot "data\norms.xlsx") -Destination (Join-Path $releaseApp "data\norms.xlsx")
Copy-Item -LiteralPath (Join-Path $projectRoot "templates\card_template.xlsx") -Destination (Join-Path $releaseApp "templates\card_template.xlsx")

$releaseOutputFiles = @(Get-ChildItem -LiteralPath (Join-Path $releaseApp "output") -File -Recurse)
$releaseRequestFiles = @(Get-ChildItem -LiteralPath (Join-Path $releaseApp "requests") -File -Recurse)
if ($releaseOutputFiles.Count -ne 0 -or $releaseRequestFiles.Count -ne 0) {
    throw "Unexpected files were copied into release working directories."
}

Write-Output "Portable build created: $releaseApp"
