$ErrorActionPreference = "Stop"

$materialRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$sourceRoot = Join-Path $materialRoot "源码"
$pythonExe = Join-Path $materialRoot "便携Python运行环境\python.exe"
$testTemp = Join-Path $materialRoot "测试临时目录"

if (-not (Test-Path -LiteralPath $pythonExe)) {
    throw "未找到便携Python：$pythonExe"
}
if (-not (Test-Path -LiteralPath (Join-Path $sourceRoot "tests"))) {
    throw "未找到自动测试目录。"
}

New-Item -ItemType Directory -Force -Path $testTemp | Out-Null
$env:TEMP = $testTemp
$env:TMP = $testTemp
$env:PYTHONUTF8 = "1"
$env:PYTHONIOENCODING = "utf-8"
$env:PYTHONDONTWRITEBYTECODE = "1"

Push-Location $sourceRoot
try {
    & $pythonExe -m unittest discover -s tests -v
    if ($LASTEXITCODE -ne 0) {
        throw "自动测试失败，退出码：$LASTEXITCODE"
    }
}
finally {
    Pop-Location
}

Write-Host "全部自动测试通过。" -ForegroundColor Green
