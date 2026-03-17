from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from zhihuishu_pro_ppt_downloader.manifest import (
    download_manifest_resources,
    dump_json,
    export_manifest_as_individual_pdfs,
    export_manifest_as_merged_pdf,
    load_manifest,
    reconcile_manifest_with_downloads,
    summarize_manifest,
)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="zhihuishu-pro-ppt-downloader",
        description="Audit and download Zhihuishu Pro PPT manifests exported by the Tampermonkey script.",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    summary_parser = subparsers.add_parser("manifest-summary", help="Print a summary for a resource manifest JSON file.")
    summary_parser.add_argument("manifest", help="Path to the exported resource manifest JSON.")
    summary_parser.add_argument("--json", action="store_true", help="Emit JSON instead of a human-readable summary.")

    reconcile_parser = subparsers.add_parser(
        "manifest-reconcile",
        help="Compare a resource manifest against a downloads directory.",
    )
    reconcile_parser.add_argument("manifest", help="Path to the exported resource manifest JSON.")
    reconcile_parser.add_argument(
        "--downloads-dir",
        default="downloads",
        help="Directory containing downloaded PPT/PPTX files. Defaults to ./downloads",
    )
    reconcile_parser.add_argument("--json", action="store_true", help="Emit JSON instead of a human-readable summary.")
    reconcile_parser.add_argument("--output", help="Optional path to save the reconciliation JSON report.")

    download_parser = subparsers.add_parser(
        "manifest-download",
        help="Automatically download all PPT resources in the manifest.",
    )
    download_parser.add_argument("manifest", help="Path to the exported resource manifest JSON.")
    download_parser.add_argument(
        "--downloads-dir",
        default="downloads",
        help="Directory to save the PPT files. Defaults to ./downloads",
    )
    download_parser.add_argument("--no-skip", action="store_false", dest="skip_existing", help="Redownload existing files.")

    pdf_parser = subparsers.add_parser(
        "manifest-export-pdf",
        help="Convert PPTs to PDF and merge them into one file with bookmarks (Requires PowerPoint on Windows).",
    )
    pdf_parser.add_argument("manifest", help="Path to the exported resource manifest JSON.")
    pdf_parser.add_argument(
        "--downloads-dir",
        default="downloads",
        help="Directory containing the PPT files. Defaults to ./downloads",
    )
    pdf_parser.add_argument(
        "--output",
        default="course_merged.pdf",
        help="Output path for the merged PDF. Defaults to course_merged.pdf",
    )
    pdf_parser.add_argument(
        "--individual",
        action="store_true",
        help="Export each PPT as an individual PDF file instead of merging.",
    )
    pdf_parser.add_argument(
        "--output-dir",
        default="pdfs",
        help="Output directory for individual PDFs. Defaults to ./pdfs",
    )

    return parser


def render_summary(summary: dict) -> str:
    course = summary.get("course", {})
    stats = summary.get("stats", {})
    lines = [
        f"课程: {course.get('mapName') or '未知课程'}",
        f"课程 ID: {course.get('courseId') or ''}",
        f"班级 ID: {course.get('classId') or ''}",
        f"Manifest 生成时间: {summary.get('generatedAt') or ''}",
        f"扫描知识点: {stats.get('scannedNodes', 0)} / {stats.get('totalNodes', 0)}",
        f"资源总数: {stats.get('totalResources', 0)}",
        f"PPT 节点数: {summary.get('pptNodeCount', 0)}",
        f"PPT 资源数: {summary.get('pptResourceCount', 0)}",
        f"零计数但实际有资源的节点: {stats.get('zeroCountButActualResourcesNodes', 0)}",
        f"零计数但实际有 PPT 的节点: {stats.get('zeroCountButPptNodes', 0)}",
        f"接口错误节点数: {stats.get('errorCount', 0)}",
        f"PPT 后缀分布: {summary.get('suffixCounter')}",
        f"PPT 资源类型分布: {summary.get('resourceTypeCounter')}",
        f"PPT 数据类型分布: {summary.get('resourceDataTypeCounter')}",
    ]
    duplicates = summary.get("duplicateExpectedFilenames") or []
    if duplicates:
        lines.append(f"重复期望文件名: {duplicates}")
    dup_urls = summary.get("duplicateDownloadUrls") or []
    if dup_urls:
        lines.append(f"重复下载链接: {dup_urls}")
    return "\n".join(lines)


def render_reconcile(report: dict) -> str:
    course = report.get("course", {})
    lines = [
        f"课程: {course.get('mapName') or '未知课程'}",
        f"下载目录: {report.get('downloadsDir') or ''}",
        f"期望 PPT 文件数: {report.get('expectedPptFiles', 0)}",
        f"实际 PPT 文件数: {report.get('actualPptFiles', 0)}",
        f"匹配文件数: {report.get('matchedFiles', 0)}",
        f"缺失文件数: {report.get('missingFiles', 0)}",
        f"多余文件数: {report.get('unexpectedFiles', 0)}",
        f"重复实际文件名数: {report.get('duplicateActualFiles', 0)}",
    ]

    missing = report.get("missing") or []
    if missing:
        lines.append("缺失文件:")
        lines.extend(f"  - {item['expectedFilename']} ({item['knowledgeName']})" for item in missing[:20])
        if len(missing) > 20:
            lines.append(f"  - ... 还有 {len(missing) - 20} 个")

    unexpected = report.get("unexpected") or []
    if unexpected:
        lines.append("多余文件:")
        lines.extend(f"  - {item}" for item in unexpected[:20])
        if len(unexpected) > 20:
            lines.append(f"  - ... 还有 {len(unexpected) - 20} 个")

    duplicates = report.get("duplicateActual") or {}
    if duplicates:
        lines.append("重复实际文件:")
        for name, paths in list(duplicates.items())[:20]:
            lines.append(f"  - {name}: {len(paths)} 份")
        if len(duplicates) > 20:
            lines.append(f"  - ... 还有 {len(duplicates) - 20} 组")

    return "\n".join(lines)


def render_download(results: dict) -> str:
    lines = [
        "下载完成！",
        f"总资源数: {results['total']}",
        f"新下载数: {results['downloaded']}",
        f"跳过已存在数: {results['skipped']}",
        f"失败数: {results['failed']}",
    ]
    if results["errors"]:
        lines.append("错误详情:")
        for err in results["errors"][:10]:
            lines.append(f"  - {err['filename']}: {err['error']}")
        if len(results["errors"]) > 10:
            lines.append(f"  - ... 还有 {len(results['errors']) - 10} 个错误")
    return "\n".join(lines)


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    manifest = load_manifest(args.manifest)

    if args.command == "manifest-summary":
        summary = summarize_manifest(manifest)
        if args.json:
            print(json.dumps(summary, ensure_ascii=False, indent=2))
        else:
            print(render_summary(summary))
        return 0

    if args.command == "manifest-reconcile":
        report = reconcile_manifest_with_downloads(manifest, args.downloads_dir)
        if args.output:
            output_path = dump_json(report, args.output)
            print(f"已写入: {output_path}")
        if args.json:
            print(json.dumps(report, ensure_ascii=False, indent=2))
        else:
            print(render_reconcile(report))
        return 0

    if args.command == "manifest-download":
        results = download_manifest_resources(manifest, args.downloads_dir, args.skip_existing)
        print(render_download(results))
        return 0 if results["failed"] == 0 else 1

    if args.command == "manifest-export-pdf":
        try:
            if args.individual:
                output_path = export_manifest_as_individual_pdfs(manifest, args.downloads_dir, args.output_dir)
                print(f"\n[OK] 导出成功！个别 PDF 已保存至: {output_path}")
            else:
                output_path = export_manifest_as_merged_pdf(manifest, args.downloads_dir, args.output)
                print(f"\n[OK] 导出成功！合并 PDF 已保存至: {output_path}")
            return 0
        except Exception as e:
            print(f"\n[FAIL] 导出失败: {e}")
            return 1

    parser.error(f"Unknown command: {args.command}")
    return 2


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
