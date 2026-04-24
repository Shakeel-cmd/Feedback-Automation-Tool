import zipfile
import shutil
from pathlib import Path


def create_zip_from_folder(source_folder: Path, output_path: Path) -> bytes:
    """
    Creates a ZIP archive of source_folder at output_path.
    Returns the ZIP content as bytes and deletes the temp file.
    """
    shutil.make_archive(str(output_path).replace(".zip", ""), "zip", source_folder)
    with open(output_path, "rb") as f:
        data = f.read()
    output_path.unlink(missing_ok=True)
    return data


def create_lob_zip(session_folder: Path, lob_name: str, output_path: Path) -> bytes:
    """
    Creates a ZIP with excel/{lob}/ and pdf/{lob}/ structure for a specific LOB.
    Returns the ZIP content as bytes and deletes the temp file.
    """
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for source_dir, prefix in [
            (session_folder / "excel" / lob_name, f"excel/{lob_name}"),
            (session_folder / "pdf"   / lob_name, f"pdf/{lob_name}"),
        ]:
            if source_dir.exists():
                for file in source_dir.iterdir():
                    if file.is_file():
                        zf.write(file, f"{prefix}/{file.name}")
    with open(output_path, "rb") as f:
        data = f.read()
    output_path.unlink(missing_ok=True)
    return data
