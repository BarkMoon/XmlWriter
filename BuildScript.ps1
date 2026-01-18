$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
$msbuildPath = $null

if (Test-Path $vswhere) {
    $msbuildPath = & $vswhere -latest -products * -requires Microsoft.Component.MSBuild -find MSBuild\**\Bin\MSBuild.exe
}

if (-not $msbuildPath -or -not (Test-Path $msbuildPath)) {
    Write-Host "VS MSBuild not found. Trying .NET Framework MSBuild..." -ForegroundColor Yellow
    $msbuildPath = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\MSBuild.exe"
}

if (-not (Test-Path $msbuildPath)) {
    Write-Error "MSBuild not found."
    exit 1
}

$msbuildPath = $msbuildPath.Trim()
Write-Host "Using MSBuild: $msbuildPath"

$logDir = "Log"
if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir | Out-Null
}
$logFile = Join-Path $logDir "build_log.txt"

& $msbuildPath "XmlWriter.csproj" /t:Build /p:Configuration=Release /v:m | Tee-Object -FilePath $logFile
