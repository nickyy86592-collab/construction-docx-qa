$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$cacheRoot = $env:RUNNER_TOOL_CACHE
if (-not $cacheRoot) { $cacheRoot = Join-Path $env:USERPROFILE '.cache\construction-docx-qa' }
$dest = Join-Path $cacheRoot 'changfeng-audit-python312'
New-Item -ItemType Directory -Force $dest | Out-Null
$python = Join-Path $dest 'python.exe'
if (-not (Test-Path $python)) {
  Write-Host 'Preparing isolated Python 3.12.10 for original hard audit scripts'
  $zip = Join-Path $dest 'python.zip'
  Invoke-WebRequest -UseBasicParsing -Uri 'https://www.python.org/ftp/python/3.12.10/python-3.12.10-embed-amd64.zip' -OutFile $zip
  Expand-Archive -LiteralPath $zip -DestinationPath $dest -Force
}
$site = Join-Path $dest 'Lib\site-packages'
if (-not (Test-Path (Join-Path $site 'win32com\__init__.py'))) {
  Write-Host 'Preparing pywin32 wheel for matching CPython version'
  $meta = Invoke-RestMethod -Uri 'https://pypi.org/pypi/pywin32/json'
  $wheel = $meta.urls | Where-Object { $_.filename -match 'cp312-cp312-win_amd64.whl$' } | Select-Object -First 1
  if ($null -eq $wheel) { throw 'Matching pywin32 wheel not found' }
  $wheelZip = Join-Path $dest 'pywin32.zip'
  Invoke-WebRequest -UseBasicParsing -Uri $wheel.url -OutFile $wheelZip
  $hash = (Get-FileHash $wheelZip -Algorithm SHA256).Hash.ToLower()
  if ($hash -ne $wheel.digests.sha256) { throw 'pywin32 wheel integrity check failed' }
  New-Item -ItemType Directory -Force $site | Out-Null
  Expand-Archive -LiteralPath $wheelZip -DestinationPath $site -Force
  Copy-Item (Join-Path $site 'pywin32_system32\*.dll') -Destination $dest -Force
}
@('python312.zip','.','Lib\site-packages','Lib\site-packages\win32','Lib\site-packages\win32\lib','Lib\site-packages\pywin32_system32','import site') | Set-Content -Encoding ASCII (Join-Path $dest 'python312._pth')
if (-not (Test-Path (Join-Path $site 'PIL\__init__.py'))) {
  Write-Host 'Preparing cached Pillow for embedded color-image checks'
  $meta = Invoke-RestMethod -Uri 'https://pypi.org/pypi/Pillow/json'
  $wheel = $meta.urls | Where-Object { $_.filename -match 'cp312-cp312-win_amd64.whl$' } | Select-Object -First 1
  if ($null -eq $wheel) { throw 'Matching Pillow wheel not found' }
  $wheelZip = Join-Path $dest 'pillow.zip'
  Invoke-WebRequest -UseBasicParsing -Uri $wheel.url -OutFile $wheelZip
  $hash = (Get-FileHash $wheelZip -Algorithm SHA256).Hash.ToLower()
  if ($hash -ne $wheel.digests.sha256) { throw 'Pillow wheel integrity check failed' }
  Expand-Archive -LiteralPath $wheelZip -DestinationPath $site -Force
}
& $python -c "import sys,win32com.client; from PIL import Image; print(sys.version); print('PYWIN32_AND_PILLOW_READY')"
if ($LASTEXITCODE -ne 0) { throw 'Isolated Python / pywin32 validation failed' }
$env:AUDIT_PYTHON = $python
"AUDIT_PYTHON=$python" | Out-File -FilePath $env:GITHUB_ENV -Encoding utf8 -Append

