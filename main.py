"""
main.py — NEU Curriculum (CTĐT) Extractor

Lệnh:
  python main.py --test              # Test 5 file đầu, in báo cáo QA
  python main.py                     # Toàn bộ file → MinIO qbcur/
  python main.py --local ./output    # Lưu JSON local
  python main.py --doc ./doccur      # Xử lý thư mục local (không cần MinIO)
  python main.py --dry-run           # Chỉ list tên file
  python main.py --no-skip           # Overwrite file đã tồn tại
"""
import argparse
import json
import logging
import re
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone
from pathlib import Path

import config
from parser import parse_curriculum
from storage import MinIOStore

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler("curriculum_extractor.log", encoding="utf-8"),
    ],
)
log = logging.getLogger(__name__)


# ── Worker ────────────────────────────────────────────────────────────────────

def process_one(store, object_name, skip, local_dir):
    filename = store.filename_of(object_name)

    if skip and local_dir is None and store.already_processed(object_name):
        return {"status": "skip", "file": filename, "qa": None}

    try:
        file_bytes = store.download(object_name)
        record     = parse_curriculum(file_bytes, filename)
        data       = record.to_dict()
        qa         = data["_qa"]
        data["_meta"]["source_file"] = filename

        if local_dir:
            safe = re.sub(r"\.docx$", ".json", filename, flags=re.IGNORECASE)
            out  = local_dir / safe
            out.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
            dest = str(out)
        else:
            dest = store.upload_json(object_name, data)

        icon = "✓" if qa["is_ok"] else "⚠"
        log.info(
            f"[{icon}] {record.major_code or '???':15s} | "
            f"TC={record.total_credits or '?':>4} | "
            f"PLO={len(record.plos):2d} | "
            f"HP={len(record.courses):3d} | "
            f"score={qa['completeness_score']:3.0f} | "
            f"{filename[:55]}"
        )
        for issue in (qa["issues"] or []):
            log.warning(f"      ↳ {issue}")

        return {
            "status": "ok" if qa["is_ok"] else "warn",
            "file": filename,
            "code": record.major_code,
            "qa": qa,
            "dest": dest,
        }

    except Exception as exc:
        log.error(f"[✗] {filename}: {exc}", exc_info=True)
        return {
            "status": "error",
            "file": filename,
            "code": "",
            "qa": {
                "is_ok": False,
                "issues": [f"PARSE ERROR: {exc}"],
                "completeness_score": 0,
                "needs_review": True,
            },
        }


# ── QA Report ─────────────────────────────────────────────────────────────────

def print_qa_report(results: list, test_mode: bool = False):
    total   = len(results)
    ok      = sum(1 for r in results if r["status"] == "ok")
    warn    = sum(1 for r in results if r["status"] == "warn")
    errors  = sum(1 for r in results if r["status"] == "error")
    skipped = sum(1 for r in results if r["status"] == "skip")

    print("\n" + "═" * 65)
    print(f"  BÁO CÁO QA CTĐT {'(TEST MODE)' if test_mode else ''}")
    print("═" * 65)
    print(f"  Tổng:         {total}")
    print(f"  ✓ Đạt:        {ok}")
    print(f"  ⚠ Cần review: {warn}")
    print(f"  ✗ Lỗi parse: {errors}")
    print(f"  → Bỏ qua:    {skipped}")

    problem_files = [r for r in results if r["status"] in ("warn", "error")]
    if problem_files:
        print(f"\n  ── File cần xem lại ({len(problem_files)}) ──────────────")
        for r in problem_files:
            score = r["qa"]["completeness_score"] if r["qa"] else 0
            print(f"\n  [{r['status'].upper()}] {r['file']}")
            print(f"         Score: {score}/100")
            if r["qa"] and r["qa"]["issues"]:
                for issue in r["qa"]["issues"]:
                    print(f"         • {issue}")
    else:
        print("\n  ✓ Không có file nào cần xem lại!")
    print("═" * 65)

    ts = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    report = {
        "generated_at": ts,
        "summary": {"total": total, "ok": ok, "warn": warn, "errors": errors, "skipped": skipped},
        "problem_files": [
            {"file": r["file"], "code": r.get("code", ""), "qa": r["qa"]}
            for r in problem_files
        ],
    }
    rp = Path(f"qa_curriculum_{ts}.json")
    rp.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\n  Report lưu tại: {rp}")


# ── Doc-folder mode ───────────────────────────────────────────────────────────

def _run_doc_mode(args):
    doc_dir = Path(args.doc)
    out_dir = Path(args.local) if args.local else Path("./output_curriculum")
    out_dir.mkdir(parents=True, exist_ok=True)

    files = sorted(doc_dir.glob("*.docx"))
    if not files:
        log.warning(f"Không tìm thấy .docx nào trong {doc_dir}")
        return

    log.info(f"Doc mode: {len(files)} file từ {doc_dir.resolve()} → {out_dir.resolve()}")

    results = []
    for f in files:
        filename = f.name
        try:
            file_bytes = f.read_bytes()
            record     = parse_curriculum(file_bytes, filename)
            data       = record.to_dict()
            qa         = data["_qa"]
            data["_meta"]["source_file"] = filename

            safe = re.sub(r"\.docx$", ".json", filename, flags=re.IGNORECASE)
            out  = out_dir / safe
            out.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

            icon = "✓" if qa["is_ok"] else "⚠"
            log.info(
                f"[{icon}] {record.major_code or '???':<15s} | "
                f"TC={record.total_credits or '?':>4} | "
                f"PLO={len(record.plos):2d} | "
                f"HP={len(record.courses):3d} | "
                f"score={qa['completeness_score']:3.0f} | "
                f"{filename[:55]}"
            )
            for issue in qa["issues"]:
                log.warning(f"      ↳ {issue}")

            results.append({
                "status": "ok" if qa["is_ok"] else "warn",
                "file": filename,
                "code": record.major_code,
                "qa": qa,
                "dest": str(out),
            })

        except Exception as exc:
            log.error(f"[✗] {filename}: {exc}", exc_info=True)
            results.append({
                "status": "error", "file": filename, "code": "",
                "qa": {"is_ok": False, "issues": [f"PARSE ERROR: {exc}"],
                       "completeness_score": 0, "needs_review": True},
            })

    print_qa_report(results, test_mode=True)


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    ap = argparse.ArgumentParser(description="NEU Curriculum (CTĐT) Extractor")
    ap.add_argument("--test",    action="store_true", help="Test 5 file đầu + báo cáo QA")
    ap.add_argument("--limit",   type=int, default=0, help="Giới hạn số file (0=tất cả)")
    ap.add_argument("--no-skip", action="store_true", help="Overwrite file đã xử lý")
    ap.add_argument("--dry-run", action="store_true", help="Chỉ list file")
    ap.add_argument("--local",   type=str, default=None, help="Lưu JSON local thay vì MinIO")
    ap.add_argument("--doc",     type=str, default=None, help="Thư mục .docx local (bỏ qua MinIO)")
    ap.add_argument("--workers", type=int, default=config.MAX_WORKERS)
    args = ap.parse_args()

    if args.test:
        args.limit = 5
        if not args.local:
            args.local = "./output_curriculum"

    if args.doc:
        return _run_doc_mode(args)

    store = MinIOStore(
        endpoint=config.MINIO_ENDPOINT,
        access_key=config.MINIO_ACCESS_KEY,
        secret_key=config.MINIO_SECRET_KEY,
        bucket=config.MINIO_BUCKET,
        input_prefix=config.MINIO_INPUT_PREFIX,
        output_prefix=config.MINIO_OUTPUT_PREFIX,
        secure=config.MINIO_SECURE,
    )

    objects = store.list_docx()
    if args.limit:
        objects = objects[:args.limit]

    log.info(
        f"Input:  {config.MINIO_BUCKET}/{config.MINIO_INPUT_PREFIX}\n"
        f"         Output: {config.MINIO_BUCKET}/{config.MINIO_OUTPUT_PREFIX}\n"
        f"         Files:  {len(objects)} | workers={args.workers}"
    )

    if args.dry_run:
        for o in objects: print(o)
        return

    local_dir = None
    if args.local:
        local_dir = Path(args.local)
        local_dir.mkdir(parents=True, exist_ok=True)

    skip = not args.no_skip
    results = []

    with ThreadPoolExecutor(max_workers=args.workers) as pool:
        futures = {
            pool.submit(process_one, store, obj, skip, local_dir): obj
            for obj in objects
        }
        for fut in as_completed(futures):
            results.append(fut.result())

    print_qa_report(results, test_mode=args.test)


if __name__ == "__main__":
    main()
