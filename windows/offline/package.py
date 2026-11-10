"""Assemble the Windows x64 installer from its compiled stubs and offline inputs."""

import argparse
import hashlib
import shutil
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZIP_DEFLATED, ZIP_STORED, ZipFile

OFFICE_NAME = "LibreOffice-Win-x64.msi"
OFFICE_HASH = "480fecf759fc6393bd614fbf43753eb95a8198fc1a065d41b727e3b664cd1e59"
WHEELS = (
    "lxml-6.1.1-cp312-cp312-win_amd64.whl",
    "pillow-12.3.0-cp312-cp312-win_amd64.whl",
    "pypdfium2-5.14.0-py3-none-win_amd64.whl",
)


def setup_runtime(runtime, nuget, tcl, wheels):
    with ZipFile(nuget) as package:
        for entry in package.infolist():
            parts = entry.filename.split("/")
            if parts[0] != "tools" or entry.is_dir() or parts[1] in {"include", "libs"}:
                continue
            target = runtime.joinpath(*parts[1:])
            target.parent.mkdir(parents=True, exist_ok=True)
            with package.open(entry) as source, target.open("wb") as output:
                shutil.copyfileobj(source, output)
    for folder in ("DLLs", "Lib/tkinter", "tcl"):
        shutil.copytree(tcl / folder, runtime / folder, dirs_exist_ok=True)
    site_packages = runtime / "Lib/site-packages"
    for name in WHEELS:
        with ZipFile(wheels / name) as wheel:
            wheel.extractall(site_packages)


def package(args):
    with args.libreoffice.open("rb") as stream:
        if hashlib.file_digest(stream, "sha256").hexdigest() != OFFICE_HASH:
            raise ValueError("O MSI não corresponde ao LibreOffice 26.2.7 x64 publicado.")
    project = Path(__file__).resolve().parents[2]
    with TemporaryDirectory() as temporary:
        app = Path(temporary) / "app"
        app.mkdir()
        shutil.copyfile(args.launcher, app / "Tributos-Docs.exe")
        shutil.copyfile(project / "app.py", app / "app.py")
        shutil.copytree(project / "tributos", app / "tributos",
                        ignore=shutil.ignore_patterns("__pycache__", "*.pyc"))
        shutil.copyfile(Path(__file__).with_name("TERCEIROS.txt"), app / "TERCEIROS.txt")
        setup_runtime(app / "runtime", args.python_nuget, args.tcl, args.wheels)
        args.output.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(args.stub, args.output)
        with ZipFile(args.output, "a", ZIP_DEFLATED, compresslevel=6) as archive:
            for path in sorted(app.rglob("*")):
                if path.is_file():
                    archive.write(path, "app/" + path.relative_to(app).as_posix())
            archive.write(args.libreoffice, OFFICE_NAME, compress_type=ZIP_STORED)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    for argument in ("stub", "launcher", "python-nuget", "tcl", "wheels", "libreoffice", "output"):
        parser.add_argument("--" + argument, required=True, type=Path)
    package(parser.parse_args())
