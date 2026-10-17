"""Checks the update helpers that do not need a packaged app."""
import hashlib
import os
import sys
import tempfile

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))

import updater


def run():
    data = b"new version"
    asset = {"name": updater.ASSET_NAMES[updater.platform_key()], "size": len(data),
             "digest": "sha256:" + hashlib.sha256(data).hexdigest()}
    assert updater.find_asset({"assets": [{"name": "other"}, asset]}) is asset
    assert updater.find_asset({"assets": [{"name": "other"}]}) is None

    updater.verify(data, asset)
    for bad in (b"new versioN", b"short"):
        try:
            updater.verify(bad, asset)
        except ValueError:
            pass
        else:
            raise AssertionError("corrupted download was accepted")

    # Running from source never tries to replace itself
    assert not updater.can_self_update()

    # Swap a staged file into place, the way an executable is replaced
    folder = tempfile.mkdtemp()
    target = os.path.join(folder, "ARHerb")
    with open(target, "w") as f:
        f.write("old")
    staging = tempfile.mkdtemp(prefix=updater.STAGING_PREFIX, dir=folder)
    staged = os.path.join(staging, "ARHerb")
    with open(staged, "w") as f:
        f.write("new")
    updater.install(staged, target)
    with open(target) as f:
        assert f.read() == "new"
    assert sorted(os.listdir(folder)) == ["ARHerb"], os.listdir(folder)
    print("updater test passed")


if __name__ == "__main__":
    run()
