# DeepSeek 办公助手 - 本地 HTTPS 证书生成与信任
# 输出: cert\localhost.pfx（服务端证书+私钥）、cert\localhost.cer（公钥）
# 信任: 将证书加入当前用户「受信任的根证书颁发机构」（无需管理员权限）
# 幂等: 可重复运行；证书存在则只做信任检查
$ErrorActionPreference = 'Stop'
$dir = Split-Path -Parent $MyInvocation.MyCommand.Path
$certDir = Join-Path $dir 'cert'
if (-not (Test-Path $certDir)) { New-Item -ItemType Directory -Path $certDir | Out-Null }
$pfxPath = Join-Path $certDir 'localhost.pfx'
$cerPath = Join-Path $certDir 'localhost.cer'
$pass = 'deepseek-dev'

# 1) 生成（缺任一文件才重新生成）
if (-not (Test-Path $pfxPath) -or -not (Test-Path $cerPath)) {
    Write-Host '生成自签名证书 localhost（SAN: localhost, 127.0.0.1，有效期 10 年）...'
    $cert = New-SelfSignedCertificate `
        -DnsName 'localhost','127.0.0.1' `
        -CertStoreLocation 'Cert:\CurrentUser\My' `
        -FriendlyName 'DeepSeek Office Add-in localhost' `
        -NotAfter (Get-Date).AddYears(10) `
        -KeyExportPolicy Exportable `
        -KeyAlgorithm RSA -KeyLength 2048 `
        -Type SSLServerAuthentication
    Export-PfxCertificate -Cert $cert -FilePath $pfxPath `
        -Password (ConvertTo-SecureString -String $pass -Force -AsPlainText) | Out-Null
    Export-Certificate -Cert $cert -FilePath $cerPath | Out-Null
    Write-Host "已导出证书: $pfxPath"
}

# 2) 信任（按指纹去重，重复运行不会重复导入）
$thumb = (New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($cerPath)).Thumbprint
$existing = Get-ChildItem Cert:\CurrentUser\Root | Where-Object { $_.Thumbprint -eq $thumb }
if ($existing) {
    Write-Host '证书已在「受信任的根证书颁发机构」中，跳过导入。'
} else {
    try {
        # certutil 无 UI 弹窗，非交互环境下也能快速完成或快速失败
        $out = & certutil.exe -user -addstore Root $cerPath 2>&1
        if ($LASTEXITCODE -ne 0) { throw ($out -join ' ') }
        Write-Host '已将证书加入当前用户「受信任的根证书颁发机构」。'
    } catch {
        Write-Warning ('信任导入失败: ' + $_.Exception.Message)
        Write-Warning '证书文件已生成，但未被信任。请重新运行本脚本，或手动将 localhost.cer 导入「受信任的根证书颁发机构」。'
        exit 1
    }
}
Write-Host '证书就绪。'
exit 0
