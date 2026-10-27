param(
    [switch]$Force
)

$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$tool = Join-Path $root '源码\通道生效签名短链工具'
$runtimeBase = Join-Path $root '便携Python运行环境'
$build = Join-Path $root '构建临时目录'
$runtime = Join-Path $build 'runtime'
$release = Join-Path $build '发布'
$payload = Join-Path $build 'runtime_payload.zip'
$resource = Join-Path $root '发布资源'
$icon = Join-Path $resource '应用图标.ico'
$manifest = Join-Path $tool '应用清单.manifest'
$antd = Join-Path $root '源码\第三方\AntdUI\AntdUI.dll'
$csc = 'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe'
$makensis = Join-Path $root '构建依赖\NSIS-3.12\makensis.exe'
$versionFile = Join-Path $root '版本.json'
$resultFile = Join-Path $root '输出\构建结果.txt'

# 业务源码 -> 运行时载荷中的文件名。这个映射也必须参与缓存版本计算。
$sourceMap = [ordered]@{
    '通道生效签名短链工具源码.py' = 'app.py'
    'table_split_merge.py' = 'table_split_merge.py'
    'signature_consistency.py' = 'signature_consistency.py'
    'port_allocator.py' = 'port_allocator.py'
    'mapping_manager.py' = 'mapping_manager.py'
}

function Read-ProjectVersion {
    if (-not (Test-Path -LiteralPath $versionFile)) {
        throw "缺少版本文件：$versionFile"
    }
    $json = [IO.File]::ReadAllText($versionFile, [Text.Encoding]::UTF8) | ConvertFrom-Json
    $productVersion = [string]$json.productVersion
    $releaseDateText = [string]$json.releaseDate
    if ($productVersion -notmatch '^\d+\.\d+\.\d+$') {
        throw "版本.json 的 productVersion 必须是 主版本.次版本.修订号，例如 1.4.0；当前为：$productVersion"
    }
    try {
        $releaseDate = [datetime]::ParseExact($releaseDateText, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
    }
    catch {
        throw "版本.json 的 releaseDate 必须是 yyyy-MM-dd；当前为：$releaseDateText"
    }
    return [PSCustomObject]@{
        ProductVersion = $productVersion
        ReleaseDate = $releaseDate
        ReleaseDateText = $releaseDateText
        ReleaseCompact = $releaseDate.ToString('yyyyMMdd')
    }
}

function Get-FileSha256([string]$Path) {
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()
}

# 运行环境缓存版本由真正进入载荷的内容决定：
# Python 业务源码 + 随包模板/地区库/主码号种子 + 应用图标。
# 任何一项变化都会得到新的缓存目录，无需再手工提升 RuntimeVersion。
function Get-PayloadFingerprint {
    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($sourceName in $sourceMap.Keys) {
        $path = Join-Path $tool $sourceName
        $parts.Add("源码/$sourceName=$(Get-FileSha256 $path)")
    }
    foreach ($file in Get-ChildItem -LiteralPath $resource -File |
        Where-Object { $_.Extension -in @('.xlsx', '.png') } |
        Sort-Object Name) {
        $parts.Add("发布资源/$($file.Name)=$(Get-FileSha256 $file.FullName)")
    }
    $bytes = [Text.Encoding]::UTF8.GetBytes(($parts -join "`n"))
    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        return ([BitConverter]::ToString($sha.ComputeHash($bytes)) -replace '-', '').ToLowerInvariant()
    }
    finally {
        $sha.Dispose()
    }
}

function Write-AssemblyVersionSource([string]$Path, $Version, [string]$RuntimeVersion, [string]$MappingSeedVersion) {
    $content = @"
// 本文件由 生成安装包.ps1 自动生成，请勿手工修改。
using System.Reflection;

[assembly: AssemblyVersion("$($Version.ProductVersion).0")]
[assembly: AssemblyFileVersion("$($Version.ProductVersion).0")]
[assembly: AssemblyInformationalVersion("$($Version.ProductVersion)")]
[assembly: AssemblyTitle("通道生效签名短链程序")]
[assembly: AssemblyProduct("通道生效签名短链程序")]
[assembly: AssemblyDescription("通道生效签名短链程序")]
[assembly: AssemblyMetadata("ProductVersion", "$($Version.ProductVersion)")]
[assembly: AssemblyMetadata("ReleaseDate", "$($Version.ReleaseDateText)")]
[assembly: AssemblyMetadata("RuntimeVersion", "$RuntimeVersion")]
[assembly: AssemblyMetadata("MappingSeedVersion", "$MappingSeedVersion")]
"@
    $utf8Bom = New-Object System.Text.UTF8Encoding($true)
    [IO.File]::WriteAllText($Path, $content, $utf8Bom)
}

$version = Read-ProjectVersion
$runtimeVersion = "缓存版本待计算"
$mappingSeedVersion = "映射种子待计算"

$installerName = "通道生效签名短链程序_安装版_$($version.ProductVersion)_$($version.ReleaseCompact).exe"
$installer = Join-Path $root "输出\$installerName"
$installerRelative = "输出\$installerName"

foreach ($needed in @(
    $runtimeBase, $resource, $icon, $manifest, $antd, $csc, $makensis,
    $versionFile, (Join-Path $root '安装程序.nsi'), (Join-Path $root '运行测试.ps1')
)) {
    if (-not (Test-Path -LiteralPath $needed)) {
        throw "缺少构建文件：$needed"
    }
}
foreach ($sourceName in $sourceMap.Keys) {
    $sourcePath = Join-Path $tool $sourceName
    if (-not (Test-Path -LiteralPath $sourcePath)) {
        throw "缺少业务源码：$sourcePath"
    }
}
foreach ($resourceName in @(
    '地区库.xlsx', '子端口分配.xlsx', '最终报备表模板.xlsx',
    '通道生效签名导入模板.xlsx', '通道生效签名短链导入模板.xlsx',
    '通道编号及对应主码号.xlsx', '应用图标.png'
)) {
    $resourcePath = Join-Path $resource $resourceName
    if (-not (Test-Path -LiteralPath $resourcePath)) {
        throw "缺少发布资源：$resourcePath"
    }
}

if ((Test-Path -LiteralPath $installer) -and -not $Force) {
    throw "安装包已存在：$installer`n确认覆盖请执行：.\生成安装包.ps1 -Force"
}

$payloadFingerprint = Get-PayloadFingerprint
$runtimeVersion = "$($version.ReleaseCompact)-$($payloadFingerprint.Substring(0, 8))"
$mappingSeedVersion = "mapping-$((Get-FileSha256 (Join-Path $resource '通道编号及对应主码号.xlsx')).Substring(0, 8))"

Write-Host "产品版本：$($version.ProductVersion)（$($version.ReleaseDateText)）" -ForegroundColor Cyan
Write-Host "运行环境版本：$runtimeVersion" -ForegroundColor Cyan
Write-Host "主码号映射种子：$mappingSeedVersion" -ForegroundColor Cyan
Write-Host "安装包：$installerName" -ForegroundColor Cyan

Write-Host '运行自动测试……' -ForegroundColor Cyan
& (Join-Path $root '运行测试.ps1')
if ($LASTEXITCODE -ne 0) { throw '自动测试失败，停止打包。' }
$testStatus = '全部通过'

$rootFull = [IO.Path]::GetFullPath($root).TrimEnd('\') + '\'
$buildFull = [IO.Path]::GetFullPath($build)
if (-not $buildFull.StartsWith($rootFull, [StringComparison]::OrdinalIgnoreCase)) { throw '构建目录不在开发目录内。' }
if (Test-Path -LiteralPath $buildFull) { Remove-Item -LiteralPath $buildFull -Recurse -Force }
New-Item -ItemType Directory -Force -Path $build, $release | Out-Null
Copy-Item -LiteralPath $runtimeBase -Destination $runtime -Recurse -Force

foreach ($entry in $sourceMap.GetEnumerator()) {
    Copy-Item -LiteralPath (Join-Path $tool $entry.Key) -Destination (Join-Path $runtime $entry.Value) -Force
}

$internalResources = Join-Path $runtime 'resources'
New-Item -ItemType Directory -Force -Path $internalResources | Out-Null
Get-ChildItem -LiteralPath $resource -File -Filter '*.xlsx' | Copy-Item -Destination $internalResources -Force
Copy-Item -LiteralPath (Join-Path $resource '应用图标.png') -Destination $internalResources -Force
Compress-Archive -Path (Join-Path $runtime '*') -DestinationPath $payload -CompressionLevel Optimal -Force

$assemblyInfo = Join-Path $build 'AssemblyVersion.g.cs'
Write-AssemblyVersionSource $assemblyInfo $version $runtimeVersion $mappingSeedVersion

Write-Host '编译 AntdUI 窗口和内置运行环境……' -ForegroundColor Cyan
& $csc /nologo /target:winexe /optimize+ /platform:x64 "/out:$release\通道生效签名短链程序.exe" "/win32icon:$icon" "/win32manifest:$manifest" "/resource:$payload,runtime_payload.zip" `
    /reference:System.Windows.Forms.dll /reference:System.Drawing.dll /reference:System.IO.Compression.dll `
    /reference:System.Web.Extensions.dll "/reference:$antd" `
    (Join-Path $tool '便携启动器源码.cs') (Join-Path $tool '主码号管理源码.cs') $assemblyInfo
if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath (Join-Path $release '通道生效签名短链程序.exe'))) {
    throw '窗口编译失败。'
}
Copy-Item -LiteralPath $antd -Destination $release -Force

Write-Host '生成安装程序……' -ForegroundColor Cyan
Push-Location $root
try {
    & $makensis /INPUTCHARSET UTF8 /V2 "/DAPPVERSION=$($version.ProductVersion)" "/DOUTFILE=$installerRelative" '安装程序.nsi'
    if ($LASTEXITCODE -ne 0) { throw 'NSIS 安装包生成失败。' }
}
finally {
    Pop-Location
}
if (-not (Test-Path -LiteralPath $installer)) { throw '安装包文件不存在。' }

$installerHash = Get-FileSha256 $installer
$installerSizeMb = [math]::Round((Get-Item -LiteralPath $installer).Length / 1MB, 2)
$gitCommit = ''
try { $gitCommit = (& git -C $root rev-parse --short HEAD 2>$null) } catch { }
$gitDirty = ''
try { if (& git -C $root status --porcelain 2>$null) { $gitDirty = '（含未提交修改）' } } catch { }

New-Item -ItemType Directory -Force -Path (Split-Path -Parent $resultFile) | Out-Null
$resultLines = @(
    '通道生效签名短链程序 构建结果',
    '',
    "产品版本：$($version.ProductVersion)",
    "发布日期：$($version.ReleaseDateText)",
    "运行环境版本：$runtimeVersion",
    "主码号映射种子：$mappingSeedVersion",
    "构建时间：$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')",
    "Git 提交：$gitCommit$gitDirty",
    "自动测试：$testStatus",
    '',
    "安装包：$installer",
    "安装包大小：$installerSizeMb MB",
    "SHA-256：$installerHash"
)
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
[IO.File]::WriteAllLines($resultFile, $resultLines, $utf8Bom)

Write-Host "安装包：$installer" -ForegroundColor Green
Write-Host "安装包大小：$installerSizeMb MB" -ForegroundColor Green
Write-Host "SHA-256：$installerHash" -ForegroundColor Green
Write-Host "构建完成，结果清单：$resultFile" -ForegroundColor Green
