$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $PSScriptRoot
$Dist = Join-Path $Root "dist"
$Stage = Join-Path $Dist "remote-codex-feishu-bridge"
$Zip = Join-Path $Dist "remote-codex-feishu-bridge.zip"

if (Test-Path $Stage) {
  Remove-Item $Stage -Recurse -Force
}
New-Item -ItemType Directory -Path $Stage | Out-Null
New-Item -ItemType Directory -Path $Dist -Force | Out-Null

$includeFiles = @(
  "README.md",
  "INSTALL.md",
  "package.json",
  "package-lock.json",
  ".env.example",
  ".gitignore",
  "feishu-agent.js",
  "start-feishu.cmd",
  "install.cmd"
)

foreach ($file in $includeFiles) {
  Copy-Item (Join-Path $Root $file) (Join-Path $Stage $file) -Force
}

New-Item -ItemType Directory -Path (Join-Path $Stage "scripts") | Out-Null
Copy-Item (Join-Path $Root "scripts\make-release.ps1") (Join-Path $Stage "scripts\make-release.ps1") -Force

if (Test-Path $Zip) {
  Remove-Item $Zip -Force
}
Compress-Archive -Path (Join-Path $Stage "*") -DestinationPath $Zip -Force

Write-Host "Release package created:"
Write-Host $Zip
