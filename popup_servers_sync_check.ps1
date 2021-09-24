Add-Type -AssemblyName System.Windows.Forms
$global:balmsg = New-Object System.Windows.Forms.NotifyIcon
$path = (Get-Process -id $pid).Path
$balmsg.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path)
$balmsg.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Warning
$balmsg.BalloonTipText = 'Время проверить синхронизацию серверов по аптекам'
$balmsg.BalloonTipTitle = "Внимание $Env:USERNAME"
$balmsg.Visible = $true
$balmsg.ShowBalloonTip(60000)