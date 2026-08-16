# DeepSeek 办公助手 - 注册到本机 Word / Excel（微软官方开发侧载机制）
# 与 office-addin-dev-settings 的 registerAddIn 一致（Word/Excel 路径）：
#   HKCU\Software\Microsoft\Office\16.0\Wef\Developer 下
#   写 值名=插件ID, 值=manifest.xml 绝对路径；Office 下次启动时加载。
# 卸载: 删除该注册表值即可（或重装本脚本覆盖）。
$ErrorActionPreference = 'Stop'
$dir = Split-Path -Parent $MyInvocation.MyCommand.Path
$m = Join-Path $dir 'manifest.xml'
$x = [xml](Get-Content -LiteralPath $m -Raw -Encoding UTF8)
$id = $x.OfficeApp.Id
$key = 'HKCU:\Software\Microsoft\Office\16.0\Wef\Developer'
New-Item -Path $key -Force | Out-Null
New-ItemProperty -Path $key -Name $id -Value $m -PropertyType String -Force | Out-Null
# 清理旧版脚本可能误写的 Outlook 专用子键条目（Word/Excel 不读那里）
$oldKey = 'HKCU:\Software\Microsoft\Office\16.0\Wef\Developer\OutlookSideloadManifestPath'
if (Test-Path $oldKey) { Remove-ItemProperty -Path $oldKey -Name $id -ErrorAction SilentlyContinue }
Write-Host "已注册插件 ID: $id"
Write-Host "清单路径: $m"
Write-Host '完全退出并重新打开 Word / Excel 后生效（开始选项卡出现「DeepSeek 助手」分组）。'
exit 0
