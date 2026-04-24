import os
import zipfile
import shutil
from pathlib import Path


def create_zip_from_folder(source_folder: Path, output_path: Path) -> bytes:
    """
    Creates a ZIP archive of source_folder at output_path.
    Returns the ZIP content as bytes and deletes the temp file.
    """
    print(f"📦 ZIP — files being added:")
    if source_folder.exists():
        for f in sorted(source_folder.rglob("*")):
            if f.is_file():
                print(f"   {f} — exists: {os.path.exists(f)}")
    else:
        print(f"   (source folder does not exist: {source_folder})")
    shutil.make_archive(str(output_path).replace(".zip", ""), "zip", source_folder)
    print(f"📦 ZIP saved at: {output_path}")
    print(f"📦 ZIP size: {os.path.getsize(output_path)} bytes")
    with open(output_path, "rb") as f:
        data = f.read()
    output_path.unlink(missing_ok=True)
    return data


def create_lob_zip(session_folder: Path, lob_name: str, output_path: Path) -> bytes:
    """
    Creates a ZIP with excel/{lob}/ and pdf/{lob}/ structure for a specific LOB.
    Returns the ZIP content as bytes and deletes the temp file.
    """
    print(f"📦 ZIP — files being added:")
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for source_dir, prefix in [
            (session_folder / "excel" / lob_name, f"excel/{lob_name}"),
            (session_folder / "pdf"   / lob_name, f"pdf/{lob_name}"),
        ]:
            if source_dir.exists():
                for file in source_dir.iterdir():
                    if file.is_file():
                        print(f"   {file} — exists: {os.path.exists(file)}")
                        zf.write(file, f"{prefix}/{file.name}")
            else:
                print(f"   (folder missing: {source_dir})")
    print(f"📦 ZIP saved at: {output_path}")
    print(f"📦 ZIP size: {os.path.getsize(output_path)} bytes")
    with open(output_path, "rb") as f:
        data = f.read()
    output_path.unlink(missing_ok=True)
    return data
