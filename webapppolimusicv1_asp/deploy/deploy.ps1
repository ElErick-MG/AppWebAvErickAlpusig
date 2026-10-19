# =============================================================================
#  PoliMusic - IIS deployment script
#  Copies the application into C:\inetpub\wwwroot, creates a dedicated
#  application pool ("No Managed Code") and verifies that the site answers.
#
#  Run it from an elevated PowerShell console:
#      powershell -ExecutionPolicy Bypass -File .\deploy.ps1
#  or simply double-click deploy.bat (it asks for elevation).
# =============================================================================
[CmdletBinding()]
param(
    [string]$Source      = "",
    [string]$Destination = "C:\inetpub\wwwroot",
    [string]$AppPool     = "PoliMusic",
    [string]$SiteName    = "Default Web Site",
    [switch]$SkipCopy,
    [switch]$SkipIisConfig
)

$ErrorActionPreference = "Stop"

function Write-Step($text) {
    Write-Host ""
    Write-Host "==> $text" -ForegroundColor Cyan
}

function Write-Ok($text) {
    Write-Host "    [OK] $text" -ForegroundColor Green
}

function Write-Warn($text) {
    Write-Host "    [!!] $text" -ForegroundColor Yellow
}

# -----------------------------------------------------------------------------
# 0. Elevation check
# -----------------------------------------------------------------------------
$identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($identity)
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "This script must run as Administrator." -ForegroundColor Red
    Write-Host "Right-click deploy.bat and choose 'Run as administrator'." -ForegroundColor Red
    exit 1
}

if ([string]::IsNullOrWhiteSpace($Source)) {
    $Source = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
}

if (-not (Test-Path (Join-Path $Source "login.asp"))) {
    Write-Host "Source folder does not look like the PoliMusic app: $Source" -ForegroundColor Red
    exit 1
}

Write-Host "PoliMusic deployment" -ForegroundColor White
Write-Host "  source      : $Source"
Write-Host "  destination : $Destination"
Write-Host "  app pool    : $AppPool"

# -----------------------------------------------------------------------------
# 1. Required IIS features
# -----------------------------------------------------------------------------
Write-Step "Checking IIS features"
Import-Module WebAdministration -ErrorAction Stop

$features = @("IIS-WebServer", "IIS-StaticContent", "IIS-DefaultDocument", "IIS-ASP", "IIS-ISAPIExtensions")
foreach ($f in $features) {
    try {
        $state = (Get-WindowsOptionalFeature -Online -FeatureName $f -ErrorAction Stop).State
        if ($state -ne "Enabled") {
            Write-Warn "$f is disabled - enabling it (this can take a minute)"
            Enable-WindowsOptionalFeature -Online -FeatureName $f -All -NoRestart | Out-Null
        }
        else {
            Write-Ok "$f enabled"
        }
    }
    catch {
        Write-Warn "Could not query $f : $($_.Exception.Message)"
    }
}

# -----------------------------------------------------------------------------
# 2. Copy the application files
# -----------------------------------------------------------------------------
if (-not (Test-Path $Destination)) {
    New-Item -ItemType Directory -Path $Destination -Force | Out-Null
    Write-Ok "Created $Destination"
}

if (-not $SkipCopy) {
    Write-Step "Copying application files"

    $excludeDirs = @("deploy", "docs", "SongFilesAndImages_resources", ".workbuddy-ai", ".git")
    $excludeFiles = @("*.sql", "*.py", "*.md", "*.ps1", "*.bat")

    $robocopyArgs = @(
        "`"$Source`"", "`"$Destination`"", "/E", "/R:2", "/W:1",
        "/NFL", "/NDL", "/NJH", "/NJS", "/NP"
    )
    if ($excludeDirs.Count -gt 0) {
        $robocopyArgs += "/XD"
        foreach ($d in $excludeDirs) { $robocopyArgs += "`"$Source\$d`"" }
    }
    if ($excludeFiles.Count -gt 0) {
        $robocopyArgs += "/XF"
        foreach ($f in $excludeFiles) { $robocopyArgs += "`"$f`"" }
    }

    $cmd = "robocopy " + ($robocopyArgs -join " ")
    cmd.exe /c $cmd | Out-Null
    $rc = $LASTEXITCODE
    if ($rc -ge 8) {
        Write-Host "robocopy failed with exit code $rc" -ForegroundColor Red
        exit 1
    }
    Write-Ok "Files copied (robocopy exit code $rc)"
}

# -----------------------------------------------------------------------------
# 3. Permissions for the IIS worker process
# -----------------------------------------------------------------------------
Write-Step "Granting read access to IIS_IUSRS"
$aclTarget = $Destination
& icacls.exe $aclTarget /grant "IIS_IUSRS:(OI)(CI)RX" /T /C /Q | Out-Null
if ($LASTEXITCODE -eq 0) {
    Write-Ok "ACL applied to $aclTarget"
}
else {
    Write-Warn "icacls returned $LASTEXITCODE - check permissions manually"
}

# -----------------------------------------------------------------------------
# 4-6. Application pool, binding and smoke test
#      (skipped when publishing into an existing application folder)
# -----------------------------------------------------------------------------
if ($SkipIisConfig) {
    Write-Step "Skipping IIS configuration (-SkipIisConfig)"
    Write-Ok "Files are in place; the existing application keeps its pool"
    $testUrl = "http://localhost/" + (Split-Path -Leaf $Destination) + "/login.asp"
}
else {
    Write-Step "Configuring application pool '$AppPool'"
    if (-not (Test-Path "IIS:\AppPools\$AppPool")) {
        New-WebAppPool -Name $AppPool | Out-Null
        Write-Ok "Application pool created"
    }
    else {
        Write-Ok "Application pool already exists"
    }

    Set-ItemProperty "IIS:\AppPools\$AppPool" -Name managedRuntimeVersion -Value ""
    Set-ItemProperty "IIS:\AppPools\$AppPool" -Name managedPipelineMode  -Value "Integrated"
    Set-ItemProperty "IIS:\AppPools\$AppPool" -Name enable32BitAppOnWin64 -Value $false
    Write-Ok "Pipeline: Integrated / No Managed Code / 64-bit"

    Write-Step "Binding site '$SiteName' to the application pool"
    try {
        $site = Get-Website -Name $SiteName
        if ($null -eq $site) {
            Write-Warn "Site '$SiteName' not found. Assign the pool manually in IIS Manager."
        }
        else {
            Set-ItemProperty "IIS:\Sites\$SiteName" -Name applicationPool -Value $AppPool
            Write-Ok "Site '$SiteName' now uses '$AppPool' (state: $($site.State))"
            if ($site.State -ne "Started") {
                Start-Website -Name $SiteName
                Write-Ok "Site started"
            }
        }
    }
    catch {
        Write-Warn "Could not bind the pool: $($_.Exception.Message)"
    }

    Write-Step "Recycling the application pool"
    Restart-WebAppPool $AppPool
    Start-Sleep -Seconds 2
    $testUrl = "http://localhost/login.asp"
}

Write-Step "Smoke test"
try {
    $response = Invoke-WebRequest -Uri $testUrl -UseBasicParsing -TimeoutSec 25
    if ($response.StatusCode -eq 200) {
        Write-Ok "GET $testUrl -> 200 ($($response.RawContentLength) bytes)"
    }
    else {
        Write-Warn "GET $testUrl -> $($response.StatusCode)"
    }
}
catch {
    Write-Warn "Could not reach $testUrl : $($_.Exception.Message)"
    Write-Warn "Open IIS Manager, select the site and check the ASP settings / bindings."
}

Write-Host ""
Write-Host "Deployment finished." -ForegroundColor Green
Write-Host "Open $testUrl and sign in with:"
Write-Host "  userAdmin      / 12345678   (administrator)"
Write-Host "  Erick.Alpusig  / 12345678   (listener)"
Write-Host ""
Write-Host "If the page shows a 500 error, enable detailed ASP errors:"
Write-Host "  IIS Manager > site > ASP > Debugging Properties > Send Errors To Browser = True"
Write-Host "  IIS Manager > site > Error Pages > Edit Feature Settings > Detailed errors"
