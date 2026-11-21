"""审计底稿抽取 Web 应用（Flask 工厂）。"""
from __future__ import annotations

import json
import os
import secrets
import shutil
import time
from pathlib import Path

from flask import (Flask, abort, render_template, request,
                   send_file, session)

from .engine.extractor import build_extract
from .engine.model import WorkbookLoadError, WorkbookModel
from .engine.tracer import InputError, parse_a1, parse_rect, trace


def create_app(instance_path: str | None = None) -> Flask:
    app = Flask(__name__, instance_path=instance_path)
    app.config.update(
        SECRET_KEY=os.environ.get("SECRET_KEY", secrets.token_hex(32)),
        MAX_CONTENT_LENGTH=8 * 1024 * 1024,
        RUN_RETENTION_SECONDS=7 * 24 * 3600,
    )
    runs_dir = Path(app.instance_path) / "runs"
    runs_dir.mkdir(parents=True, exist_ok=True)

    def _run_dir(run_id: str) -> Path:
        if not run_id or not run_id.replace("-", "").isalnum():
            abort(404)
        d = runs_dir / run_id
        if not d.exists():
            abort(404, "会话已过期，请重新上传文件")
        return d

    def _gc_runs():
        now = time.time()
        for d in runs_dir.iterdir():
            try:
                if d.is_dir() and now - d.stat().st_mtime > app.config["RUN_RETENTION_SECONDS"]:
                    shutil.rmtree(d, ignore_errors=True)
            except OSError:
                continue

    @app.get("/")
    def index():
        return render_template("index.html")

    @app.post("/upload")
    def upload():
        _gc_runs()
        f = request.files.get("workbook")
        if not f or not f.filename:
            return render_template("index.html", error="请选择要上传的 .xlsx 文件"), 400
        if not f.filename.lower().endswith(".xlsx"):
            return render_template("index.html", error="仅支持 .xlsx 格式"), 400
        run_id = secrets.token_hex(12)
        d = runs_dir / run_id
        d.mkdir(parents=True)
        src_path = d / "source.xlsx"
        f.save(src_path)
        try:
            model = WorkbookModel(str(src_path))
        except WorkbookLoadError as exc:
            shutil.rmtree(d, ignore_errors=True)
            return render_template("index.html", error=str(exc)), 400
        sheet_info = [
            {"name": name, "rows": model.dimensions(name)[0],
             "cols": model.dimensions(name)[1]}
            for name in model.sheet_names
        ]
        session["run_id"] = run_id
        session["sheet_names"] = model.sheet_names
        return render_template("config.html",
                               sheet_info=sheet_info,
                               filename=f.filename,
                               indicators=[{"sheet": model.sheet_names[0], "addr": ""}],
                               forbidden=[{"sheet": model.sheet_names[0], "addr": ""}])

    def _read_rows(prefix: str) -> list[dict]:
        sheets = request.form.getlist(f"{prefix}_sheet")
        addrs = request.form.getlist(f"{prefix}_addr")
        return [{"sheet": s, "addr": a} for s, a in zip(sheets, addrs)]

    @app.post("/analyze")
    def analyze():
        run_id = session.get("run_id")
        d = _run_dir(run_id)
        model = WorkbookModel(str(d / "source.xlsx"))
        known = model.sheet_names

        ind_rows = [r for r in _read_rows("ind") if r["addr"].strip()]
        fb_rows = [r for r in _read_rows("fb") if r["addr"].strip()]
        form_errors: list[str] = []
        if not ind_rows:
            form_errors.append("至少填写一个指标格")

        indicator_keys: list[tuple] = []
        for r in ind_rows:
            try:
                key = parse_a1(r["addr"], known, default_sheet=r["sheet"])
            except InputError as exc:
                form_errors.append(f"指标格 {r['sheet']}!{r['addr']}：{exc}")
                continue
            if key not in indicator_keys:
                indicator_keys.append(key)
        if not indicator_keys:
            form_errors.append("至少填写一个有效的指标格")

        forbidden_rects: list[tuple] = []
        for r in fb_rows:
            raw = r["addr"] if "!" in r["addr"] else f"{r['sheet']}!{r['addr']}"
            try:
                rect = parse_rect(raw, known)
            except InputError as exc:
                form_errors.append(f"禁发区域 {r['sheet']}!{r['addr']}：{exc}")
                continue
            forbidden_rects.append(rect)

        if form_errors:
            sheet_info = [
                {"name": n, "rows": model.dimensions(n)[0], "cols": model.dimensions(n)[1]}
                for n in known
            ]
            return render_template("config.html", sheet_info=sheet_info,
                                   filename="source.xlsx",
                                   indicators=ind_rows or [{"sheet": known[0], "addr": ""}],
                                   forbidden=fb_rows or [{"sheet": known[0], "addr": ""}],
                                   error="；".join(form_errors)), 400

        result = trace(model, indicator_keys, forbidden_rects)
        (d / "trace.json").write_text(
            json.dumps({
                "indicators": [list(k) for k in result.indicators],
                "required": [list(k) for k in sorted(result.required)],
                "errors": [
                    {"indicator": e.indicator_text, "kind": e.kind,
                     "message": e.message, "chain": e.chain_text}
                    for e in result.errors
                ],
            }, ensure_ascii=False, default=str),
            encoding="utf-8",
        )

        if result.errors:
            return render_template("result.html", result=result, ok=False,
                                   layout=None, extract=None)

        out_path = d / "audit_extract.xlsx"
        extract = build_extract(model, result, str(out_path))
        if extract.verify_errors:
            # 核对未通过：删除产物，保证下载端点也无法取得文件
            out_path.unlink(missing_ok=True)
            return render_template("result.html", result=result, ok=False,
                                   layout=None, extract=extract,
                                   fatal="复算核对未通过，已阻止下载："
                                         + "；".join(extract.verify_errors))
        return render_template("result.html", result=result, ok=True,
                               layout=extract.sheet_layout, extract=extract)

    @app.get("/download")
    def download():
        run_id = session.get("run_id")
        d = _run_dir(run_id)
        out = d / "audit_extract.xlsx"
        if not out.exists():
            abort(404, "尚无可下载文件（追溯存在错误或未完成分析）")
        return send_file(out, as_attachment=True,
                         download_name="audit_extract.xlsx")

    @app.get("/healthz")
    def healthz():
        return {"status": "ok"}

    @app.errorhandler(413)
    def too_large(_e):
        return render_template("index.html",
                               error="上传文件超过 8 MB 限制"), 413

    return app
