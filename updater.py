"""Self-update of the packaged app (PyInstaller build) from GitHub Releases.

The new version is downloaded in the background, unpacked next to the
installed app and swapped in when the app quits (or on "Restart now").
"""
import glob
import hashlib
import os
import shutil
import subprocess
import sys
import tarfile
import tempfile

# Must match the asset names produced by .github/workflows/build.yml
ASSET_NAMES = {
    "windows": "ARHerb-windows.exe",
    "macos": "ARHerb-macos-arm64.zip",
    "linux": "ARHerb-linux-x86_64.tar.gz",
}
STAGING_PREFIX = ".arherb-update-"


def platform_key():
    if sys.platform.startswith("win"):
        return "windows"
    if sys.platform == "darwin":
        return "macos"
    if sys.platform.startswith("linux"):
        return "linux"
    return None


def install_target():
    """What gets replaced: the .app bundle on macOS, the executable elsewhere."""
    executable = os.path.realpath(sys.executable)
    if platform_key() == "macos":
        # .../ARHerb.app/Contents/MacOS/ARHerb
        return os.path.dirname(os.path.dirname(os.path.dirname(executable)))
    return executable


def can_self_update():
    if not getattr(sys, "frozen", False) or platform_key() is None:
        return False
    target = install_target()
    if platform_key() == "macos" and (not target.endswith(".app") or "/AppTranslocation/" in target):
        # An app that was not moved out of Downloads runs from a read-only copy
        return False
    return os.access(os.path.dirname(target), os.W_OK)


def find_asset(release):
    name = ASSET_NAMES.get(platform_key())
    for asset in release.get("assets", []):
        if asset.get("name") == name:
            return asset
    return None


def verify(data, asset):
    if len(data) != asset.get("size"):
        raise ValueError("Downloaded file has an unexpected size")
    digest = asset.get("digest") or ""
    if digest.startswith("sha256:") and hashlib.sha256(data).hexdigest() != digest.split(":", 1)[1]:
        raise ValueError("Downloaded file is corrupted (checksum mismatch)")


def stage(data, target):
    """Unpack the downloaded asset next to target and return the unpacked path."""
    staging = tempfile.mkdtemp(prefix=STAGING_PREFIX, dir=os.path.dirname(target))
    try:
        key = platform_key()
        if key == "macos":
            archive = os.path.join(staging, "update.zip")
            with open(archive, "wb") as f:
                f.write(data)
            # ditto keeps the symlinks and permissions inside the bundle
            subprocess.run(["ditto", "-x", "-k", archive, staging], check=True)
            os.remove(archive)
            staged = os.path.join(staging, "ARHerb.app")
        elif key == "linux":
            archive = os.path.join(staging, "update.tar.gz")
            with open(archive, "wb") as f:
                f.write(data)
            with tarfile.open(archive) as tar:
                tar.extract(tar.getmember("ARHerb"), staging, filter="data")
            os.remove(archive)
            staged = os.path.join(staging, "ARHerb")
            os.chmod(staged, 0o755)
        else:
            staged = os.path.join(staging, "ARHerb.exe")
            with open(staged, "wb") as f:
                f.write(data)
        if not os.path.exists(staged):
            raise FileNotFoundError("The update archive does not contain ARHerb")
        return staged
    except Exception:
        shutil.rmtree(staging, ignore_errors=True)
        raise


def remove_path(path):
    if os.path.isdir(path) and not os.path.islink(path):
        shutil.rmtree(path, ignore_errors=True)
    elif os.path.lexists(path):
        try:
            os.remove(path)
        except OSError:
            pass  # Windows: still locked by the old process, retried on next launch


def install(staged, target):
    old = target + ".old"
    remove_path(old)
    # Windows allows renaming a running .exe; macOS/Linux keep running from the moved files
    os.replace(target, old)
    try:
        os.replace(staged, target)
    except OSError:
        os.replace(old, target)
        raise
    shutil.rmtree(os.path.dirname(staged), ignore_errors=True)
    remove_path(old)


def relaunch(target):
    env = dict(os.environ)
    # Start a fresh PyInstaller instance instead of reusing this one's unpacked files
    env["PYINSTALLER_RESET_ENVIRONMENT"] = "1"
    if platform_key() == "macos":
        subprocess.Popen(["open", "-n", target], env=env)
    else:
        subprocess.Popen([target], env=env, close_fds=True)


def cleanup_leftovers():
    """Remove what a previous update left behind (old version, unfinished downloads)."""
    if not getattr(sys, "frozen", False) or platform_key() is None:
        return
    target = install_target()
    remove_path(target + ".old")
    for staging in glob.glob(os.path.join(os.path.dirname(target), STAGING_PREFIX + "*")):
        shutil.rmtree(staging, ignore_errors=True)
