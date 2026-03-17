from __future__ import annotations

import json
import os
import shutil
import tempfile
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any

import httpx
from tqdm import tqdm

try:
    import comtypes.client
    HAS_COMTYPES = True
except ImportError:
    HAS_COMTYPES = False

try:
    from pypdf import PdfWriter
    HAS_PYPDF = True
except ImportError:
    HAS_PYPDF = False

PPT_SUFFIXES = {"ppt", "pptx"}


def load_manifest(path: str | Path) -> dict[str, Any]:
    manifest_path = Path(path).expanduser().resolve()
    with manifest_path.open("r", encoding="utf-8") as fh:
        return json.load(fh)


def sanitize_filename(name: str) -> str:
    return "".join("_" if ch in '<>:"/\\|?*' else ch for ch in str(name)).strip() or "downloaded_ppt"


def build_expected_filename(node: dict[str, Any], resource: dict[str, Any]) -> str:
    parts = [
        node.get("themeName") or "",
        node.get("subThemeName") or "",
        node.get("knowledgeName") or "",
        resource.get("resourcesName") or "",
    ]
    base_name = sanitize_filename(" - ".join(part for part in parts if part))
    suffix = str(resource.get("resourcesSuffix") or "").lower()
    if base_name.lower().endswith((".ppt", ".pptx")):
        return base_name
    if suffix in PPT_SUFFIXES:
        return f"{base_name}.{suffix}"
    return base_name


def iter_manifest_resources(manifest: dict[str, Any]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for node in manifest.get("nodes", []):
        for resource in node.get("resources", []):
            rows.append({"node": node, "resource": resource})
    return rows


def iter_manifest_ppts(manifest: dict[str, Any]) -> list[dict[str, Any]]:
    rows = []
    for item in iter_manifest_resources(manifest):
        resource = item["resource"]
        if resource.get("isPpt"):
            rows.append(item)
    return rows


def summarize_manifest(manifest: dict[str, Any]) -> dict[str, Any]:
    rows = iter_manifest_ppts(manifest)
    suffix_counter = Counter()
    type_counter = Counter()
    data_type_counter = Counter()
    host_counter = Counter()
    duplicate_by_name = Counter()
    duplicate_by_url = Counter()

    expected_filenames: list[str] = []
    for item in rows:
        node = item["node"]
        resource = item["resource"]
        expected_name = build_expected_filename(node, resource)
        expected_filenames.append(expected_name)
        duplicate_by_name[expected_name] += 1
        duplicate_by_url[str(resource.get("downloadUrl") or "")] += 1
        suffix_counter[str(resource.get("resourcesSuffix") or "")] += 1
        type_counter[str(resource.get("resourcesType") or "")] += 1
        data_type_counter[str(resource.get("resourcesDataType") or "")] += 1
        host_counter[str(resource.get("fromFileDomain") or "")] += 1

    nodes_with_ppt = [node for node in manifest.get("nodes", []) if int(node.get("pptCount") or 0) > 0]

    return {
        "generatedAt": manifest.get("generatedAt"),
        "course": manifest.get("course", {}),
        "stats": manifest.get("stats", {}),
        "pptNodeCount": len(nodes_with_ppt),
        "pptResourceCount": len(rows),
        "suffixCounter": dict(sorted(suffix_counter.items())),
        "resourceTypeCounter": dict(sorted(type_counter.items())),
        "resourceDataTypeCounter": dict(sorted(data_type_counter.items())),
        "fromFileDomainCounter": dict(sorted(host_counter.items())),
        "duplicateExpectedFilenames": sorted(name for name, count in duplicate_by_name.items() if count > 1),
        "duplicateDownloadUrls": sorted(url for url, count in duplicate_by_url.items() if url and count > 1),
    }


def reconcile_manifest_with_downloads(
    manifest: dict[str, Any],
    downloads_dir: str | Path,
) -> dict[str, Any]:
    downloads_path = Path(downloads_dir).expanduser().resolve()
    actual_files = sorted(path for path in downloads_path.rglob("*") if path.is_file() and path.suffix.lower() in {".ppt", ".pptx"})
    actual_by_name: dict[str, list[Path]] = defaultdict(list)
    for path in actual_files:
        actual_by_name[path.name].append(path)

    expected_entries = []
    for item in iter_manifest_ppts(manifest):
        node = item["node"]
        resource = item["resource"]
        expected_filename = build_expected_filename(node, resource)
        matches = actual_by_name.get(expected_filename, [])
        expected_entries.append(
            {
                "knowledgeName": node.get("knowledgeName") or "",
                "nodeUid": node.get("nodeUid") or "",
                "expectedFilename": expected_filename,
                "resourcesName": resource.get("resourcesName") or "",
                "resourcesUid": resource.get("resourcesUid") or "",
                "downloadUrl": resource.get("downloadUrl") or "",
                "matched": bool(matches),
                "matchedPaths": [str(path) for path in matches],
            }
        )

    missing = [entry for entry in expected_entries if not entry["matched"]]
    expected_name_set = {entry["expectedFilename"] for entry in expected_entries}
    unexpected = [str(path) for path in actual_files if path.name not in expected_name_set]
    duplicate_actual = {
        name: [str(path) for path in paths]
        for name, paths in sorted(actual_by_name.items())
        if len(paths) > 1
    }

    return {
        "course": manifest.get("course", {}),
        "manifestStats": manifest.get("stats", {}),
        "downloadsDir": str(downloads_path),
        "expectedPptFiles": len(expected_entries),
        "actualPptFiles": len(actual_files),
        "matchedFiles": len(expected_entries) - len(missing),
        "missingFiles": len(missing),
        "unexpectedFiles": len(unexpected),
        "duplicateActualFiles": len(duplicate_actual),
        "missing": missing,
        "unexpected": unexpected,
        "duplicateActual": duplicate_actual,
    }


def dump_json(data: dict[str, Any], output_path: str | Path) -> Path:
    output = Path(output_path).expanduser().resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    return output


def download_manifest_resources(
    manifest: dict[str, Any],
    downloads_dir: str | Path,
    skip_existing: bool = True,
) -> dict[str, Any]:
    """Download all PPT resources in the manifest to the specified directory."""
    downloads_path = Path(downloads_dir).expanduser().resolve()
    downloads_path.mkdir(parents=True, exist_ok=True)

    ppts = iter_manifest_ppts(manifest)
    results = {
        "total": len(ppts),
        "downloaded": 0,
        "skipped": 0,
        "failed": 0,
        "errors": [],
    }

    if not ppts:
        return results

    # Standard headers to mimic a browser request
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Referer": "https://ai-smart-course-student-pro.zhihuishu.com/",
    }

    with httpx.Client(headers=headers, follow_redirects=True, timeout=30.0) as client:
        for item in tqdm(ppts, desc="下载 PPT 进度", unit="file"):
            node = item["node"]
            resource = item["resource"]
            url = resource.get("downloadUrl")
            if not url:
                continue

            filename = build_expected_filename(node, resource)
            file_path = downloads_path / filename

            if skip_existing and file_path.exists():
                results["skipped"] += 1
                continue

            try:
                with client.stream("GET", url) as response:
                    if response.status_code != 200:
                        raise httpx.HTTPStatusError(
                            f"HTTP {response.status_code}",
                            request=response.request,
                            response=response,
                        )

                    total_size = int(response.headers.get("Content-Length", 0))

                    with file_path.open("wb") as f:
                        with tqdm(
                            total=total_size,
                            unit="B",
                            unit_scale=True,
                            desc=filename[:30],
                            leave=False,
                        ) as pbar:
                            for chunk in response.iter_bytes(chunk_size=8192):
                                f.write(chunk)
                                pbar.update(len(chunk))

                results["downloaded"] += 1
            except Exception as e:
                results["failed"] += 1
                results["errors"].append({"filename": filename, "url": url, "error": str(e)})
                if file_path.exists():
                    file_path.unlink()  # Cleanup partial download

    return results


def _convert_ppt_to_pdf(powerpoint: Any, ppt_path: Path, pdf_path: Path) -> bool:
    """Helper to convert a single PPT file to PDF using PowerPoint."""
    try:
        deck = powerpoint.Presentations.Open(str(ppt_path), WithWindow=False, ReadOnly=True)
        deck.SaveAs(str(pdf_path), 32)  # 32 = ppSaveAsPDF
        deck.Close()
        return True
    except Exception as e:
        print(f"\n错误: 转换 {ppt_path.name} 失败: {e}")
        return False


def export_manifest_as_individual_pdfs(
    manifest: dict[str, Any],
    downloads_dir: str | Path,
    output_dir: str | Path,
) -> Path:
    """Convert all PPTs in the manifest to individual PDF files."""
    if not HAS_COMTYPES:
        raise ImportError("Missing required library: comtypes. Please install it.")

    downloads_path = Path(downloads_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()
    output_path.mkdir(parents=True, exist_ok=True)

    ppts = iter_manifest_ppts(manifest)
    if not ppts:
        return output_path

    try:
        powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
        # powerpoint.Visible = 1
        
        pbar = tqdm(total=len(ppts), desc="个别 PDF 转换进度")
        
        for item in ppts:
            node = item["node"]
            resource = item["resource"]
            ppt_filename = build_expected_filename(node, resource)
            ppt_path = downloads_path / ppt_filename
            
            if not ppt_path.exists():
                print(f"\n警告: 未找到 PPT 文件 {ppt_filename}，跳过。")
                pbar.update(1)
                continue
            
            pdf_filename = ppt_path.stem + ".pdf"
            pdf_path = output_path / pdf_filename
            
            _convert_ppt_to_pdf(powerpoint, ppt_path, pdf_path)
            pbar.update(1)

        powerpoint.Quit()
    except Exception as e:
        print(f"\n❌ 初始化 PowerPoint 失败: {e}")
        raise

    return output_path


def export_manifest_as_merged_pdf(
    manifest: dict[str, Any],
    downloads_dir: str | Path,
    output_pdf: str | Path,
) -> Path:
    """Convert PPTs to PDF and merge them with bookmarks according to manifest order."""
    if not HAS_COMTYPES or not HAS_PYPDF:
        raise ImportError("Missing required libraries: comtypes and pypdf. Please install them.")

    downloads_path = Path(downloads_dir).expanduser().resolve()
    output_path = Path(output_pdf).expanduser().resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)

    # Temporary directory for intermediate PDFs
    temp_dir = Path(tempfile.mkdtemp(prefix="zhs_pdf_"))
    
    try:
        powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
        # powerpoint.Visible = 1 # Keep it hidden
        
        writer = PdfWriter()
        current_page = 0
        
        last_theme = None
        last_sub_theme = None
        
        nodes = manifest.get("nodes", [])
        pbar = tqdm(total=sum(node.get("pptCount", 0) for node in nodes), desc="导出 PDF 进度")
        
        for node in nodes:
            theme_name = node.get("themeName")
            sub_theme_name = node.get("subThemeName")
            knowledge_name = node.get("knowledgeName")
            
            # Bookmark logic
            if theme_name and theme_name != last_theme:
                writer.add_outline_item(theme_name, current_page)
                last_theme = theme_name
                last_sub_theme = None # Reset subtheme when theme changes
                
            if sub_theme_name and sub_theme_name != last_sub_theme:
                writer.add_outline_item(sub_theme_name, current_page, parent=None) # Level 2
                last_sub_theme = sub_theme_name
            
            knowledge_bookmark_added = False

            for resource in node.get("resources", []):
                if not resource.get("isPpt"):
                    continue
                
                ppt_filename = build_expected_filename(node, resource)
                ppt_path = downloads_path / ppt_filename
                
                if not ppt_path.exists():
                    print(f"\n警告: 未找到 PPT 文件 {ppt_filename}，跳过。")
                    pbar.update(1)
                    continue
                
                # Convert PPT to PDF
                temp_pdf_name = f"{ppt_path.stem}_{os.urandom(4).hex()}.pdf"
                temp_pdf_path = temp_dir / temp_pdf_name
                
                if _convert_ppt_to_pdf(powerpoint, ppt_path, temp_pdf_path):
                    # Merge PDF
                    with open(temp_pdf_path, "rb") as f:
                        # Add knowledge point bookmark only once for the first PPT in this node
                        if not knowledge_bookmark_added:
                            writer.add_outline_item(knowledge_name, current_page)
                            knowledge_bookmark_added = True
                            
                        # Append all pages of this PPT
                        writer.append(f)
                        
                        # Update current page count
                        from pypdf import PdfReader
                        reader = PdfReader(temp_pdf_path)
                        current_page += len(reader.pages)
                
                pbar.update(1)

        with open(output_path, "wb") as f:
            writer.write(f)
            
        powerpoint.Quit()
        return output_path
        
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)
