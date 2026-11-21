"""审计底稿抽取 Web 应用（Flask 工厂）。

存储模型：
- 每次上传对应一个 run 目录（runs/<run_id>/source.xlsx）；
- 同一份底稿上的每次“分析/生成”对应一个独立 submission 目录
  （runs/<run_id>/submissions/<submission_id>/），内含本次的 trace.json、
  meta.json 与仅属于本次的 audit_extract.xlsx，互不覆盖；
- 结果页 /result/<sub_id> 与下载 /download/<sub_id> 只认本次提交，
  下载仅在本次复算通过（meta.status == "ok"）且产物存在时开放；
- 提交 ID 为随机不可猜令牌，且下载/结果端点按目录校验，服务重启后依然有效；
- SECRET_KEY 持久化到实例目录，保证多 gunicorn worker 与重启后会话一致。
"""
from __future__ import annotations

import json
import os
import secrets
import shutil
import time
from pathlib import Path
from types import SimpleNamespace

from flask import (Flask, abort, redirect, render_template, request,
                   send_file, session, url_for)

from .engine.extractor import build_extract
from .engine.model import WorkbookLoadError, WorkbookModel
from .engine.tracer import InputError, parse_a1, parse_rect, trace

_RUN_ID_BYTES = 12          # token_hex(12) -> 24 个十六进制字符
_SUB_ID_BYTES = 16          # 更长，结果页/下载链接长期有效，不可猜
_ID_ALPHABET = set("0123456789abcdef")


def create_app(instance_path: str | None = None) -> Flask:
    app = Flask(__name__, instance_path=instance_path)
    app.config.update(
        SECRET_KEY=_load_or_create_secret_key(Path(app.instance_path)),
        MAX_CONTENT_LENGTH=8 * 1024 * 1024,
        RUN_RETENTION_SECONDS=7 * 24 * 3600,
    )
    runs_dir = Path(app.instance_path) / "runs"
    runs_dir.mkdir(parents=True, exist_ok=True)

    # ---- 路径与令牌校验 -------------------------------------------------

    def _is_id(token: str | None) -> bool:
        return bool(token) and len(token) <= 64 and all(
            ch in _ID_ALPHABET for ch in token
        )

    def _run_dir(run_id: str | None) -> Path:
        """解析当前会话所属 run 目录；不存在或非法一律 404。"""
        if not _is_id(run_id):
            abort(404)
        d = runs_dir / run_id
        if not d.is_dir() or not (d / "source.xlsx").is_file():
            abort(404, "会话已过期，请重新上传文件")
        return d

    def _find_submission(sub_id: str | None) -> tuple[Path, Path, dict]:
        """按提交 ID 在各 run 目录下定位提交，返回 (run_dir, sub_dir, meta)。

        不依赖会话：结果页链接在服务重启、其它分析进行/通过/失败后仍有效。
        """
        if not _is_id(sub_id):
            abort(404)
        for run_dir in runs_dir.iterdir():
            sub_dir = run_dir / "submissions" / sub_id
            meta_path = sub_dir / "meta.json"
            if sub_dir.is_dir() and meta_path.is_file():
                try:
                    meta = json.loads(meta_path.read_text(encoding="utf-8"))
                except (OSError, json.JSONDecodeError):
                    abort(404, "结果数据已损坏，请重新分析")
                return run_dir, sub_dir, meta
        abort(404, "找不到该次分析结果，可能已过期，请重新上传并分析")

    def _gc_runs():
        now = time.time()
        for d in runs_dir.iterdir():
            try:
                if d.is_dir() and now - d.stat().st_mtime > \
                        app.config["RUN_RETENTION_SECONDS"]:
                    shutil.rmtree(d, ignore_errors=True)
            except OSError:
                continue

    # ---- 路由 -----------------------------------------------------------

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
        run_id = secrets.token_hex(_RUN_ID_BYTES)
        d = runs_dir / run_id
        d.mkdir(parents=True)
        (d / "submissions").mkdir(parents=True)
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

    def _config_page(model: WorkbookModel, ind_rows, fb_rows, error: str):
        known = model.sheet_names
        sheet_info = [
            {"name": n, "rows": model.dimensions(n)[0], "cols": model.dimensions(n)[1]}
            for n in known
        ]
        return render_template(
            "config.html", sheet_info=sheet_info, filename="source.xlsx",
            indicators=ind_rows or [{"sheet": known[0], "addr": ""}],
            forbidden=fb_rows or [{"sheet": known[0], "addr": ""}],
            error=error), 400

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
            return _config_page(model, ind_rows, fb_rows, "；".join(form_errors))

        # 每次提交独立目录：不同次分析（含交错提交）互不覆盖、互不删除
        sub_id = secrets.token_hex(_SUB_ID_BYTES)
        sub_dir = d / "submissions" / sub_id
        sub_dir.mkdir(parents=True)
        # 刷新 run 目录 mtime，使持续使用的上传不被过期清理误删
        try:
            d.touch()
        except OSError:
            pass
        out_path = sub_dir / "audit_extract.xlsx"

        result = trace(model, indicator_keys, forbidden_rects)
        trace_payload = {
            "indicators": [list(k) for k in result.indicators],
            "required": [list(k) for k in sorted(result.required)],
            "errors": [
                {"indicator": e.indicator_text, "kind": e.kind,
                 "message": e.message, "chain": e.chain_text}
                for e in result.errors
            ],
        }
        (sub_dir / "trace.json").write_text(
            json.dumps(trace_payload, ensure_ascii=False, default=str),
            encoding="utf-8",
        )

        if result.errors:
            # 追溯失败：本次提交不产生任何产物，也不触碰其它提交
            meta = {
                "status": "trace_error",
                "created_at": time.time(),
                "fatal": None,
                "verify_errors": [],
                "has_file": False,
            }
            (sub_dir / "meta.json").write_text(
                json.dumps(meta, ensure_ascii=False), encoding="utf-8")
            return redirect(url_for("show_result", sub_id=sub_id))

        # 产物只写入本次提交目录；复算始终针对本次生成的 XLSX 与原底稿
        extract = build_extract(model, result, str(out_path))
        if extract.verify_errors:
            # 核对未通过：仅删除本次产物，其它已通过提交的文件不受影响
            out_path.unlink(missing_ok=True)
            meta = {
                "status": "verify_error",
                "created_at": time.time(),
                "fatal": "复算核对未通过，已阻止下载："
                         + "；".join(extract.verify_errors),
                "verify_errors": list(extract.verify_errors),
                "has_file": False,
                "verify": [list(row) for row in extract.verify],
                "indicator_mapping": [list(pair) for pair in
                                       extract.indicator_mapping],
                "sheet_layout": {
                    name: [list(rows), list(cols)]
                    for name, (rows, cols) in extract.sheet_layout.items()
                },
            }
            (sub_dir / "meta.json").write_text(
                json.dumps(meta, ensure_ascii=False, default=str),
                encoding="utf-8")
            return redirect(url_for("show_result", sub_id=sub_id))

        meta = {
            "status": "ok",
            "created_at": time.time(),
            "fatal": None,
            "verify_errors": [],
            "has_file": True,
            "verify": [list(row) for row in extract.verify],
            "indicator_mapping": [list(pair) for pair in
                                   extract.indicator_mapping],
            "sheet_layout": {
                name: [list(rows), list(cols)]
                for name, (rows, cols) in extract.sheet_layout.items()
            },
        }
        (sub_dir / "meta.json").write_text(
            json.dumps(meta, ensure_ascii=False, default=str),
            encoding="utf-8")
        return redirect(url_for("show_result", sub_id=sub_id))

    @app.get("/result/<sub_id>")
    def show_result(sub_id):
        _run_dir, sub_dir, meta = _find_submission(sub_id)
        try:
            trace_payload = json.loads(
                (sub_dir / "trace.json").read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            abort(404, "结果数据已损坏，请重新分析")

        trace_view = SimpleNamespace(
            errors=[SimpleNamespace(**e) for e in trace_payload["errors"]],
            indicators=trace_payload["indicators"],
        )
        ctx = {
            "sub_id": sub_id,
            "result": trace_view,
            "status": meta["status"],
            "fatal": meta.get("fatal"),
            "extract": None,
            "layout": None,
        }
        if meta["status"] == "ok" or meta["status"] == "verify_error":
            ctx["extract"] = SimpleNamespace(
                indicator_mapping=[tuple(p) for p in meta["indicator_mapping"]],
                verify=[tuple(row) for row in meta["verify"]],
                verify_errors=meta.get("verify_errors", []),
            )
            ctx["layout"] = {
                name: (tuple(rows), tuple(cols))
                for name, (rows, cols) in meta["sheet_layout"].items()
            }
        return render_template("result.html", **ctx)

    @app.get("/download/<sub_id>")
    def download(sub_id):
        _run_dir, sub_dir, meta = _find_submission(sub_id)
        # 仅本次提交复算通过且产物存在时才开放下载
        if meta.get("status") != "ok" or not meta.get("has_file"):
            abort(404, "该次分析未通过复算核对，不提供下载")
        out = sub_dir / "audit_extract.xlsx"
        if not out.is_file():
            abort(404, "审计件文件不存在或已被清理，请重新分析")
        return send_file(out, as_attachment=True,
                         download_name=f"audit_extract_{sub_id[:8]}.xlsx")

    @app.get("/healthz")
    def healthz():
        return {"status": "ok"}

    @app.errorhandler(413)
    def too_large(_e):
        return render_template("index.html",
                               error="上传文件超过 8 MB 限制"), 413

    return app


def _load_or_create_secret_key(instance_dir: Path) -> str:
    """优先使用环境变量；否则在实例目录持久化一个随机密钥。

    持久化后，多个 gunicorn worker 共用同一密钥，且服务重启后旧会话、
    已签发的结果页链接仍然有效。
    """
    env_key = os.environ.get("SECRET_KEY")
    if env_key:
        return env_key
    instance_dir.mkdir(parents=True, exist_ok=True)
    key_file = instance_dir / "secret_key"
    if key_file.is_file():
        value = key_file.read_text(encoding="utf-8").strip()
        if value:
            return value
    value = secrets.token_hex(32)
    key_file.write_text(value, encoding="utf-8")
    try:
        key_file.chmod(0o600)
    except OSError:
        pass
    return value
