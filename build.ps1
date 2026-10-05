# Build script for the typeset tool (rebuild). Run on a build machine with Python 3.8 (64-bit).
# Target OS: Windows 7 SP1 x64 and later.
$ErrorActionPreference = 'Stop'
$idx = 'https://pypi.tuna.tsinghua.edu.cn/simple'

# 1) runtime deps + PyInstaller
py -3.8 -m pip install -r requirements.txt -i $idx
py -3.8 -m pip install "pyinstaller==6.3.0" -i $idx

# 2) build one-file windowed exe
py -3.8 -m PyInstaller --noconfirm --clean build.spec

Write-Host ""
Write-Host "Done. Executable is under .\dist\ ."
Write-Host "For Windows 7 targets, ship the VC++ 2015-2019 redistributable (it provides UCRT),"
Write-Host "because Win7 has no Universal C Runtime installed by default."
