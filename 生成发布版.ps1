param(
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path

# 兼容旧入口：实际构建统一走 生成安装包.ps1，避免出现第二套打包逻辑。
& (Join-Path $root '生成安装包.ps1') -Force:$Force
