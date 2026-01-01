param(
  [string]$AppName = "NTOU_Tools",
  [string]$AppVersion = "1.0"
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Resolve-Path (Join-Path $root "..")
Set-Location $projectRoot

$iconSource = ".\messageImage_1767257219427.jpg"
$iconDir = ".\build\icons"
$iconPath = Join-Path $iconDir "$AppName.ico"
if (Test-Path $iconSource) {
  New-Item -ItemType Directory -Force -Path $iconDir | Out-Null
  if (-not (Test-Path $iconPath)) {
    python -m pip install --upgrade pillow | Out-Null
    python - <<'PY'
from pathlib import Path
from PIL import Image

source = Path("messageImage_1767257219427.jpg")
target = Path("build/icons/NTOU_Tools.ico")
img = Image.open(source).convert("RGBA")
sizes = [(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)]
img.save(target, sizes=sizes)
PY
  }
}

$dataArgs = @()
if (Test-Path ".\sample_excel.txt") {
  $dataArgs += "--add-data"
  $dataArgs += "sample_excel.txt;."
}
if (Test-Path ".\sample_data.json") {
  $dataArgs += "--add-data"
  $dataArgs += "sample_data.json;."
}

python -m pip install --upgrade pyinstaller
$versionFile = ".\build\windows_version.txt"
$normalizedVersion = $AppVersion.TrimStart("v", "V")
$parts = $normalizedVersion.Split(".")
$major = if ($parts.Length -ge 1) { $parts[0] } else { "1" }
$minor = if ($parts.Length -ge 2) { $parts[1] } else { "0" }
New-Item -ItemType Directory -Force -Path ".\build" | Out-Null
@"
VSVersionInfo(
  ffi=FixedFileInfo(
    filevers=($major),$minor,0,0,
    prodvers=($major),$minor,0,0,
    mask=0x3f,
    flags=0x0,
    OS=0x4,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
    ),
  kids=[
    StringFileInfo(
      [
      StringTable(
        '040904B0',
        [StringStruct('CompanyName', '$AppName'),
        StringStruct('FileDescription', '$AppName'),
        StringStruct('FileVersion', '$normalizedVersion'),
        StringStruct('InternalName', '$AppName'),
        StringStruct('OriginalFilename', '$AppName.exe'),
        StringStruct('ProductName', '$AppName'),
        StringStruct('ProductVersion', '$normalizedVersion')])
      ]),
    VarFileInfo([VarStruct('Translation', [1033, 1200])])
  ]
)
"@ | Set-Content -Path $versionFile -Encoding UTF8
$iconArgs = @()
if (Test-Path $iconPath) {
  $iconArgs += "--icon"
  $iconArgs += $iconPath
}
pyinstaller --noconfirm --clean --onefile --windowed --name $AppName --version-file $versionFile @iconArgs @dataArgs ".\Line_chart.py"

if (Test-Path ".\installer\Line_chart.iss") {
  & iscc ".\installer\Line_chart.iss"
}

$signtool = $env:SIGNTOOL_PATH
$pfxPath = $env:SIGN_CERT_PFX
$pfxPass = $env:SIGN_CERT_PASS

if ($signtool -and $pfxPath -and (Test-Path $signtool) -and (Test-Path $pfxPath)) {
  & $signtool sign /f $pfxPath /p $pfxPass /tr http://timestamp.digicert.com /td sha256 /fd sha256 ".\dist\$AppName.exe"
  $installerPath = ".\dist\installer\NTOU_Tools_Setup.exe"
  if (Test-Path $installerPath) {
    & $signtool sign /f $pfxPath /p $pfxPass /tr http://timestamp.digicert.com /td sha256 /fd sha256 $installerPath
  }
}
