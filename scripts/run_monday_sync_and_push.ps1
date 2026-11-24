$ErrorActionPreference = "Stop"
$RepoDir = Split-Path -Parent $PSScriptRoot
Set-Location $RepoDir

$EnvFile = Join-Path $RepoDir ".shiphero.env"
if (Test-Path $EnvFile) {
  Get-Content $EnvFile | ForEach-Object {
    if ($_ -match "^\s*([^#=]+)=(.*)$") {
      [System.Environment]::SetEnvironmentVariable($matches[1].Trim(), $matches[2].Trim(), "Process")
    }
  }
}

python (Join-Path $PSScriptRoot "sync_shiphero_monday.py")

git add shiphero_inventory.json psi.html index.html
$diff = git status --porcelain shiphero_inventory.json psi.html index.html
if ($diff) {
  $today = Get-Date -Format "yyyy-MM-dd"
  git commit -m "chore(psi): auto-sync Monday ShipHero I_01 inventory ($today)"
  git push origin main
}
