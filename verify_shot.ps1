$ErrorActionPreference='SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$dir = 'd:\ai工作区\Qoder CN IDE1\gw_rebuild'
$exe = "$dir\dist\公文一键排版助手.exe"
$q = Start-Process $exe -PassThru
Start-Sleep -Seconds 5
$b = [System.Windows.Forms.SystemInformation]::VirtualScreen
$bmp = New-Object System.Drawing.Bitmap($b.Width, $b.Height)
$gr = [System.Drawing.Graphics]::FromImage($bmp)
$gr.CopyFromScreen($b.Location, [System.Drawing.Point]::Empty, $b.Size)
$bmp.Save("$dir\verify_shot.png")
$gr.Dispose(); $bmp.Dispose()
if ($q -and -not $q.HasExited) { Stop-Process -Id $q.Id -Force }
"SHOT_SAVED=$dir\verify_shot.png size=" + (Get-Item "$dir\verify_shot.png").Length
