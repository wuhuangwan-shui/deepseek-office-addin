# DeepSeek 办公助手 - 一键安装：HTTPS 证书 + 注册到 Office + 结果校验
# 由 安装到Office.cmd 调用；也可单独用 PowerShell 运行本文件
$ErrorActionPreference = 'Stop'
$dir = Split-Path -Parent $MyInvocation.MyCommand.Path

function Fail($msg) {
    Write-Host ''
    Write-Host ("安装失败：" + $msg) -ForegroundColor Red
    exit 1
}

Write-Host '=============================================='
Write-Host '  DeepSeek Office Add-in  安装到 Word / Excel'
Write-Host '=============================================='

# [1/3] HTTPS 证书（生成 + 信任，幂等）
Write-Host ''
Write-Host '[1/3] 准备本地 HTTPS 证书...'
& "$dir\setup-certs.ps1"
if ($LASTEXITCODE -ne 0) { Fail '证书准备未通过，请查看上方原因。' }

# [2/3] 注册到 Office（微软官方开发侧载机制）
Write-Host ''
Write-Host '[2/3] 注册加载项到 Office...'
& "$dir\register.ps1"
if ($LASTEXITCODE -ne 0) { Fail '注册未完成，请查看上方原因。' }

# [3/3] 校验结果（注册表值、证书、信任状态）
Write-Host ''
Write-Host '[3/3] 校验安装结果...'
$m = Join-Path $dir 'manifest.xml'
$id = $null
try {
    $x = [xml](Get-Content -LiteralPath $m -Raw -Encoding UTF8)
    $id = $x.OfficeApp.Id
} catch {
    Fail ("无法读取 manifest.xml：" + $_.Exception.Message)
}
if (-not $id) { Fail 'manifest.xml 中未找到插件 ID。' }

$key = 'HKCU:\Software\Microsoft\Office\16.0\Wef\Developer'
$v = (Get-ItemProperty -Path $key -Name $id -ErrorAction SilentlyContinue).$id
if ($v -ne $m) {
    Fail ("注册表校验未通过：期望清单路径为 `"$m`"，实际值为 `"$v`"。请检查权限后重试。")
}
Write-Host ("注册表校验通过 OK（插件 ID: " + $id + "）") -ForegroundColor Green

$pfx = Join-Path $dir 'cert\localhost.pfx'
if (-not (Test-Path $pfx)) { Fail '证书文件缺失：cert\localhost.pfx 不存在。' }
Write-Host '证书文件存在 OK'

$cer = Join-Path $dir 'cert\localhost.cer'
$thumb = $null
try {
    $thumb = (New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($cer)).Thumbprint
} catch { Fail ("无法读取证书 cert\localhost.cer：" + $_.Exception.Message) }
$trusted = Get-ChildItem Cert:\CurrentUser\Root | Where-Object { $_.Thumbprint -eq $thumb }
if (-not $trusted) {
    Write-Host '警告：证书尚未被信任。若 Word 加载插件时提示证书错误，请重新运行本脚本。' -ForegroundColor Yellow
} else {
    Write-Host '证书信任校验通过 OK' -ForegroundColor Green
}

Write-Host ''
Write-Host '=============================================='
Write-Host '  安装完成！接下来：'
Write-Host '   1. 双击 start.cmd 启动本地服务（使用期间保持窗口开启）'
Write-Host '   2. 完全退出并重新打开 Word / Excel'
Write-Host '   3. 「开始」选项卡出现「DeepSeek 助手」分组'
Write-Host '      （备选入口：插入 - 我的加载项 - 共享文件夹）'
Write-Host '=============================================='
exit 0
