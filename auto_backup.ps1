# === 設定 ===
$projectPath = "C:\Users\komes\kintai-bot"  # ローカルのGASクローンフォルダ
$logFile     = "$projectPath\backup_log.txt"

# === ループ開始 ===
while ($true) {

    Set-Location $projectPath

    # GAS → ローカルの同期
    clasp pull | Out-File -Append $logFile

    # ローカル → GitHub
    git add .
    git commit -m "Auto backup $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" --allow-empty
    git push

    Start-Sleep -Seconds 60
}
while ($true) {
  Write-Host "Pulling GAS updates... " -ForegroundColor Cyan
  clasp pull
  Start-Sleep -Seconds 60   # ←間隔 60秒にしてるけど好きに変えてOK
}
