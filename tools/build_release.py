"""Build the FigTrans website's release files from the version in SKILL.md.

Run with --check to verify a checkout without updating files.
"""

import argparse
import hashlib
from html.parser import HTMLParser
from io import BytesIO
import json
import os
from pathlib import Path
import posixpath
import re
import sys
from tempfile import NamedTemporaryFile
from urllib.parse import unquote, urlsplit
from zipfile import BadZipFile, ZIP_STORED, ZipFile, ZipInfo


ROOT = Path(__file__).resolve().parents[1]
SKILL = "senior-figtrans"
MIRROR = f"source/{SKILL}/"
INDEX = "source/index.json"
MANIFEST = "downloads/release.json"
SUMS = "downloads/SHA256SUMS.txt"
DEMO = ("transformer-recreated.pptx", "transformer-recreated.vsdx", "pid-recreated.pptx")
LEGACY_ARCHIVE = "figtrans-v1.1.0.zip"
LEGACY_SHA256 = "17c86b74cf945efc91513f42ebbabb926e54a336a668d289dfa0e8fa313ab63b"
ARCHIVE_NAME = re.compile(r"figtrans-v\d+\.\d+\.\d+\.zip")
KIND = {".json": "json", ".md": "text", ".ps1": "powershell", ".py": "python",
        ".txt": "text", ".yaml": "yaml", ".yml": "yaml"}
IGNORED = {".DS_Store", "Thumbs.db", "__pycache__", ".pytest_cache"}


def sha256(data):
    return hashlib.sha256(data).hexdigest()


def json_bytes(data):
    return (json.dumps(data, ensure_ascii=False, indent=2) + "\n").encode("utf-8")


def read_json(path):
    return json.loads(path.read_text(encoding="utf-8"))


def source_files(folder):
    files = {}
    seen = set()
    for path in sorted(folder.rglob("*"), key=lambda p: (p.as_posix().casefold(), p.as_posix())):
        relative = path.relative_to(folder)
        if IGNORED.intersection(relative.parts) or path.suffix in {".pyc", ".pyo"}:
            continue
        if path.is_symlink():
            raise ValueError(f"Symlink in release contents: {path}")
        if not path.is_file():
            continue
        name = relative.as_posix()
        if name.casefold() in seen:
            raise ValueError(f"Case-insensitive duplicate release path: {name}")
        seen.add(name.casefold())
        data = path.read_bytes()
        # Windows checkouts may translate line endings. Preserve every other byte, including any BOM.
        if path.suffix in KIND or path.name == "LICENSE":
            data = data.replace(b"\r\n", b"\n")
        files[name] = data
    return files


def skill_version(files):
    for name in ("SKILL.md", "LICENSE", "THIRD_PARTY_NOTICES.md"):
        if not files.get(name):
            raise ValueError(f"Missing or empty {SKILL}/{name}")
    header = files["SKILL.md"].decode("utf-8-sig").split("---\n", 2)
    match = re.search(r'^  version:\s*[\'\"]?(\d+\.\d+\.\d+)[\'\"]?\s*$', header[1], re.M) if len(header) == 3 else None
    if match is None:
        raise ValueError("Read a semantic version from metadata.version in SKILL.md")
    return match.group(1)


def make_archive(files):
    buffer = BytesIO()
    with ZipFile(buffer, "w", compression=ZIP_STORED) as archive:
        for name, data in files.items():
            # Stored entries avoid compressor-version differences in published SHA-256 values.
            entry = ZipInfo(f"{SKILL}/{name}", date_time=(1980, 1, 1, 0, 0, 0))
            entry.create_system = 3
            entry.external_attr = 0o100644 << 16
            archive.writestr(entry, data)
    return buffer.getvalue()


class SiteLinks(HTMLParser):
    def __init__(self):
        super().__init__()
        self.urls = set()

    def handle_starttag(self, tag, attrs):
        self.urls.update(value for attr, value in attrs if value and attr in {"href", "src", "poster"})


def check_site(root, expected_files, archive_name, version):
    docs = root / "docs"
    homepage = (docs / "index.html").read_text(encoding="utf-8")
    links = SiteLinks()
    links.feed(homepage)
    script = (docs / "app.js").read_text(encoding="utf-8")
    js_paths = re.findall(r'''\b(?:fetch|fetchResource)\(\s*["']([^"']+)["']''', script)
    errors = []
    local_links = set()
    for url in sorted(links.urls | set(js_paths)):
        parsed = urlsplit(url)
        if parsed.scheme or parsed.netloc or not parsed.path:
            continue
        path = unquote(parsed.path)
        if path.startswith("/figtrans/"):
            path = path[len("/figtrans/"):]
        path = posixpath.normpath(path)
        if path == ".":
            path = "index.html"
        local_links.add(path)
        if path.startswith(("/", "../")) or path == ".." or "\\" in path:
            errors.append(f"Site link escapes docs: {url}")
        elif path not in expected_files and not (docs / path).is_file():
            errors.append(f"Missing site asset: {url}")
    required = {f"downloads/{item}" for item in (archive_name, *DEMO)}
    errors.extend(f"Missing homepage download: {item}" for item in sorted(required - local_links))

    public_text = {
        "docs/index.html": homepage,
        "docs/app.js": script,
        "README.md": (root / "README.md").read_text(encoding="utf-8"),
        "docs/docs/install.md": (docs / "docs/install.md").read_text(encoding="utf-8"),
    }
    for name, content in public_text.items():
        stale_archives = set(ARCHIVE_NAME.findall(content)) - {archive_name}
        stale_versions = set(re.findall(r"\bv(\d+\.\d+\.\d+)\b", content, re.I)) - {version}
        if stale_archives or stale_versions:
            errors.append(f"Stale release references in {name}: {', '.join(sorted(stale_archives | stale_versions))}")
    for name in ("README.md", "docs/docs/install.md"):
        if archive_name not in public_text[name]:
            errors.append(f"Missing current archive name in {name}: {archive_name}")
    return errors


def plan_release(root):
    docs = root / "docs"
    downloads = docs / "downloads"
    files = source_files(root / SKILL)
    version = skill_version(files)
    archive_name = f"figtrans-v{version}.zip"
    archive_url = f"downloads/{archive_name}"
    legacy = (downloads / LEGACY_ARCHIVE).read_bytes()
    if sha256(legacy) != LEGACY_SHA256:
        raise ValueError(f"Published historical archive has changed: {LEGACY_ARCHIVE}")
    if archive_name == LEGACY_ARCHIVE:
        with ZipFile(BytesIO(legacy)) as archive:
            contents = {name: archive.read(name) for name in archive.namelist() if not name.endswith("/")}
        if contents != {f"{SKILL}/{name}": data for name, data in files.items()}:
            raise ValueError("Skill source changed without a version bump; the v1.1.0 archive is immutable")
        archive_data = legacy
    else:
        archive_data = make_archive(files)

    manifest = read_json(docs / MANIFEST)
    manifest.update(version=version, download_url=archive_url, bytes=len(archive_data),
                    sha256=sha256(archive_data), file_count=len(files),
                    demo_downloads=[f"downloads/{name}" for name in DEMO])
    index = read_json(docs / INDEX)
    index.update(version=version, base_url=MIRROR, file_count=len(files), files=[
        {"path": name, "url": MIRROR + name, "bytes": len(data), "sha256": sha256(data),
         "type": KIND.get(Path(name).suffix, "text" if name == "LICENSE" else "binary")}
        for name, data in files.items()
    ])
    sums = {archive_name: manifest["sha256"]}
    for path in sorted(downloads.glob("figtrans-v*.zip")):
        if path.name != archive_name and ARCHIVE_NAME.fullmatch(path.name):
            sums[path.name] = sha256(path.read_bytes())
    for name in DEMO:
        sums[name] = sha256((downloads / name).read_bytes())
    expected = {MIRROR + name: data for name, data in files.items()}
    expected.update({archive_url: archive_data, INDEX: json_bytes(index), MANIFEST: json_bytes(manifest),
                     SUMS: "".join(f"{digest}  {name}\n" for name, digest in sums.items()).encode("ascii")})
    return manifest, expected, check_site(root, expected, archive_name, version)


def write_atomic(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = None
    try:
        with NamedTemporaryFile(dir=path.parent, delete=False) as handle:
            temporary = Path(handle.name)
            handle.write(data)
        os.replace(temporary, path)
    finally:
        if temporary is not None and temporary.exists():
            temporary.unlink()


def run(root, check):
    manifest, expected, errors = plan_release(root)
    docs = root / "docs"
    mirror = source_files(docs / MIRROR)
    stale = sorted(MIRROR + name for name in mirror if MIRROR + name not in expected)
    previous_mirror = {entry["path"]: entry["sha256"] for entry in read_json(docs / INDEX)["files"]}
    for path in stale:
        relative = path[len(MIRROR):]
        if check or previous_mirror.get(relative) != sha256(mirror[relative]):
            errors.append(f"Unexpected mirrored source: docs/{path}")

    updates = {}
    for name, data in expected.items():
        path = docs / name
        if path.is_symlink():
            errors.append(f"Symlink at release output: docs/{name}")
            continue
        if path.is_file():
            existing = path.read_bytes()
            mirrored_text = name.startswith(MIRROR) and (path.suffix in KIND or path.name == "LICENSE")
            if name in {INDEX, MANIFEST, SUMS} or mirrored_text:
                existing = existing.replace(b"\r\n", b"\n")
            if data == existing:
                continue
        if check:
            errors.append(f"Missing or stale release output: docs/{name}")
        else:
            updates[path] = data
    if errors:
        raise ValueError("\n".join(errors))

    for path in stale:
        (docs / path).unlink()
    for path, data in updates.items():
        write_atomic(path, data)
    action = "Verified" if check else "Built"
    print(f"{action} FigTrans {manifest['version']} ({manifest['file_count']} files, {manifest['bytes']} bytes)")
    print(f"SHA-256 {manifest['sha256']}")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--check", action="store_true", help="Check source, ZIP, mirror, metadata, and site links without writing.")
    try:
        run(ROOT, parser.parse_args().check)
    except (OSError, ValueError, KeyError, BadZipFile) as error:
        print(error, file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
