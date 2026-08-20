param(
  [string]$Version = ""
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$projectRoot = Split-Path -Parent $PSScriptRoot
Set-Location -LiteralPath $projectRoot

function Invoke-CheckedCommand {
  param(
    [Parameter(Mandatory = $true)][string]$Command,
    [Parameter(ValueFromRemainingArguments = $true)][string[]]$Arguments
  )

  & $Command @Arguments
  if ($LASTEXITCODE -ne 0) {
    throw "Command failed with exit code ${LASTEXITCODE}: $Command $($Arguments -join ' ')"
  }
}

$nodeCommand = Get-Command node.exe -ErrorAction Stop
$nodePath = $nodeCommand.Source
$nodePlatform = (& $nodePath -p "process.platform").Trim()
$nodeArch = (& $nodePath -p "process.arch").Trim()
$nodeVersion = (& $nodePath -p "process.version").Trim()

if ($nodePlatform -ne "win32" -or $nodeArch -ne "x64") {
  throw "Portable packaging requires a Windows x64 Node.js runtime. Current runtime: $nodePlatform/$nodeArch."
}

if ([string]::IsNullOrWhiteSpace($Version)) {
  $packageJson = Get-Content -LiteralPath (Join-Path $projectRoot "package.json") -Raw -Encoding UTF8 | ConvertFrom-Json
  $baseVersion = [string]$packageJson.version
  $Version = "$baseVersion-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
}

$safeVersion = $Version -replace '[^0-9A-Za-z._-]', '-'
$releaseRoot = Join-Path $projectRoot "release"
$stageRoot = Join-Path $releaseRoot "portable-staging-$safeVersion"
$appRoot = Join-Path $stageRoot "app"
$portableDist = Join-Path $appRoot "dist"
$runtimeRoot = Join-Path $stageRoot "runtime"
$zipName = "Client-Background-Investigator-Portable-win-x64-$safeVersion.zip"
$zipPath = Join-Path $releaseRoot $zipName
$hashPath = "$zipPath.sha256.txt"

New-Item -ItemType Directory -Force -Path $releaseRoot | Out-Null
if (Test-Path -LiteralPath $stageRoot) {
  Remove-Item -LiteralPath $stageRoot -Recurse -Force
}

try {
  Write-Host "[1/7] Running TypeScript checks..."
  Invoke-CheckedCommand "npm.cmd" "run" "lint"

  Write-Host "[2/7] Building the regular production application..."
  Invoke-CheckedCommand "npm.cmd" "run" "build"

  Write-Host "[3/7] Creating portable directory structure..."
  New-Item -ItemType Directory -Force -Path $portableDist, $runtimeRoot | Out-Null
  Copy-Item -LiteralPath (Join-Path $projectRoot "dist\index.html") -Destination $portableDist
  Copy-Item -LiteralPath (Join-Path $projectRoot "dist\assets") -Destination $portableDist -Recurse

  Write-Host "[4/7] Bundling the production server and runtime dependencies..."
  $portableServer = Join-Path $portableDist "server.cjs"
  Invoke-CheckedCommand "npx.cmd" "esbuild" "server.ts" "--bundle" "--platform=node" "--format=cjs" "--minify" "--define:process.env.NODE_ENV='production'" "--outfile=$portableServer"

  Write-Host "[5/7] Copying the embedded Node.js runtime ($nodeVersion)..."
  Copy-Item -LiteralPath $nodePath -Destination (Join-Path $runtimeRoot "node.exe")

  $nodeDirectory = Split-Path -Parent $nodePath
  $localLicenseCandidates = @(
    (Join-Path $nodeDirectory "LICENSE"),
    (Join-Path $nodeDirectory "LICENSE.txt")
  )
  $localLicense = $localLicenseCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
  $licenseDestination = Join-Path $runtimeRoot "LICENSE"
  if ($localLicense) {
    Copy-Item -LiteralPath $localLicense -Destination $licenseDestination
  } else {
    $previousPackage = Get-ChildItem -LiteralPath $releaseRoot -Filter "Client-Background-Investigator-Portable-win-x64-*.zip" -File |
      Sort-Object LastWriteTime -Descending |
      Select-Object -First 1
    if ($previousPackage) {
      Add-Type -AssemblyName System.IO.Compression.FileSystem
      $previousArchive = [System.IO.Compression.ZipFile]::OpenRead($previousPackage.FullName)
      try {
        $licenseEntry = $previousArchive.Entries | Where-Object {
          ($_.FullName -replace '\\', '/') -eq "runtime/LICENSE"
        } | Select-Object -First 1
        if ($licenseEntry) {
          [System.IO.Compression.ZipFileExtensions]::ExtractToFile($licenseEntry, $licenseDestination, $true)
          Write-Host "      Reused Node.js license from previous portable package."
        }
      } finally {
        $previousArchive.Dispose()
      }
    }

    if (!(Test-Path -LiteralPath $licenseDestination)) {
      $plainNodeVersion = $nodeVersion.TrimStart('v')
      $licenseUrl = "https://raw.githubusercontent.com/nodejs/node/v$plainNodeVersion/LICENSE"
      Write-Host "      Local/cached Node.js license not found; downloading $licenseUrl"
      Invoke-WebRequest -Uri $licenseUrl -OutFile $licenseDestination -UseBasicParsing
    }
  }

  if (!(Test-Path -LiteralPath $licenseDestination) -or (Get-Item -LiteralPath $licenseDestination).Length -lt 1000) {
    throw "The Node.js license file is missing or incomplete."
  }

  $portableTemplateRoot = Join-Path $projectRoot "portable"
  $launcherTemplate = Get-ChildItem -LiteralPath $portableTemplateRoot -Filter "*.bat" -File | Select-Object -First 1
  $readmeTemplate = Get-ChildItem -LiteralPath $portableTemplateRoot -Filter "*.txt" -File | Select-Object -First 1
  if (!$launcherTemplate -or !$readmeTemplate) {
    throw "Portable launcher or readme template is missing."
  }
  Copy-Item -LiteralPath $launcherTemplate.FullName -Destination $stageRoot
  Copy-Item -LiteralPath $readmeTemplate.FullName -Destination $stageRoot

  Write-Host "[6/7] Validating portable files..."
  $requiredFiles = @(
    (Join-Path $stageRoot $launcherTemplate.Name),
    (Join-Path $stageRoot $readmeTemplate.Name),
    (Join-Path $runtimeRoot "node.exe"),
    $licenseDestination,
    (Join-Path $portableDist "server.cjs"),
    (Join-Path $portableDist "index.html")
  )
  foreach ($requiredFile in $requiredFiles) {
    if (!(Test-Path -LiteralPath $requiredFile)) {
      throw "Required portable file is missing: $requiredFile"
    }
  }

  $serverCheck = & (Join-Path $runtimeRoot "node.exe") --check (Join-Path $portableDist "server.cjs") 2>&1
  if ($LASTEXITCODE -ne 0) {
    throw "Portable server syntax validation failed: $serverCheck"
  }

  Write-Host "[7/7] Creating ZIP and SHA-256 checksum..."
  if (Test-Path -LiteralPath $zipPath) {
    Remove-Item -LiteralPath $zipPath -Force
  }
  if (Test-Path -LiteralPath $hashPath) {
    Remove-Item -LiteralPath $hashPath -Force
  }
  Compress-Archive -Path (Join-Path $stageRoot "*") -DestinationPath $zipPath -CompressionLevel Optimal
  $hash = (Get-FileHash -LiteralPath $zipPath -Algorithm SHA256).Hash.ToLowerInvariant()
  Set-Content -LiteralPath $hashPath -Value "$hash  $zipName" -Encoding ASCII

  Write-Host ""
  Write-Host "Portable package created successfully:"
  Write-Host "  ZIP:    $zipPath"
  Write-Host "  SHA256: $hash"
  Write-Host "  Node:   $nodeVersion ($nodePlatform/$nodeArch)"
} finally {
  if (Test-Path -LiteralPath $stageRoot) {
    Remove-Item -LiteralPath $stageRoot -Recurse -Force
  }
}
