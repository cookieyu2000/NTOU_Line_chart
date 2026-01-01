#!/usr/bin/env python3
import platform
import shlex
import subprocess
import sys
import zipfile
from pathlib import Path

try:
    from PIL import Image
except ImportError:
    Image = None

APP_NAME = "LineChart"
ENTRYPOINT = "Line_chart.py"
ICON_SOURCE = "messageImage_1767257219427.jpg"
ASSET_DIR = "build_assets"


def ensure_pyinstaller():
    try:
        subprocess.run(
            [sys.executable, "-m", "PyInstaller", "--version"],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        return True
    except (OSError, subprocess.CalledProcessError):
        print("PyInstaller not found. Install with: python -m pip install pyinstaller pillow")
        return False


def build_icon(root, system_name):
    source = root / ICON_SOURCE
    if not source.exists():
        print(f"Icon image missing: {source}")
        return None
    if Image is None:
        print("Pillow not installed; skip icon generation.")
        return None

    assets = root / ASSET_DIR
    assets.mkdir(exist_ok=True)

    try:
        img = Image.open(source).convert("RGBA")
    except OSError:
        print("Failed to read icon image; skip icon generation.")
        return None

    png_path = assets / "app_icon.png"
    img.save(png_path)

    if system_name == "Windows":
        icon_path = assets / "app_icon.ico"
        img.save(icon_path, sizes=[(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)])
        return icon_path
    if system_name == "Darwin":
        icon_path = assets / "app_icon.icns"
        try:
            img.save(icon_path)
            return icon_path
        except OSError:
            print("Failed to write .icns file; skip icon generation.")
            return None
    return None


def build_pyinstaller_command(root, system_name, icon_path):
    data_sep = ";" if system_name == "Windows" else ":"
    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--windowed",
        "--name",
        APP_NAME,
        "--collect-all",
        "matplotlib",
    ]
    data_files = [ICON_SOURCE, "sample_data.json", "sample_excel.txt"]
    for filename in data_files:
        path = root / filename
        if path.exists():
            cmd += ["--add-data", f"{path}{data_sep}."]
    if system_name == "Windows":
        cmd.append("--onefile")
    if icon_path:
        cmd += ["--icon", str(icon_path)]
    cmd.append(str(root / ENTRYPOINT))
    return cmd


def zip_artifact(root, system_name):
    dist_dir = root / "dist"
    if system_name == "Windows":
        target = dist_dir / f"{APP_NAME}.exe"
        zip_path = dist_dir / f"{APP_NAME}-windows.zip"
    else:
        target = dist_dir / f"{APP_NAME}.app"
        zip_path = dist_dir / f"{APP_NAME}-macos.zip"

    if not target.exists():
        raise FileNotFoundError(f"Build artifact not found: {target}")

    if zip_path.exists():
        zip_path.unlink()

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as archive:
        if target.is_dir():
            for file_path in target.rglob("*"):
                archive.write(file_path, file_path.relative_to(target.parent))
        else:
            archive.write(target, target.name)
    return zip_path


def main():
    root = Path(__file__).resolve().parent
    system_name = platform.system()
    if system_name not in ("Windows", "Darwin"):
        print("Only Windows and macOS are supported.")
        return 2

    if not ensure_pyinstaller():
        return 2

    icon_path = build_icon(root, system_name)
    cmd = build_pyinstaller_command(root, system_name, icon_path)
    print("Running:", " ".join(shlex.quote(part) for part in cmd))
    subprocess.run(cmd, check=True, cwd=root)

    zip_path = zip_artifact(root, system_name)
    print(f"Done: {zip_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
