#!/usr/bin/env python3
"""
美田运营系统 v1.0
启动: python3 run.py  → http://localhost:8899
"""
import sys, os, json, uuid, csv, io
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import analyzer
import tag_manager

PORT = int(os.environ.get("PORT", 8899))
HOST = os.environ.get("HOST", "0.0.0.0")
BASE = Path(__file__).parent.resolve()
TEMP = Path(tempfile.gettempdir()) / "meitian_uploads"
TEMP.mkdir(parents=True, exist_ok=True)

# 存储合并结果的缓存
_MERGE_CACHE = {}
_DEDUP_CACHE = {}
_SUBTRACT_CACHE = {}


class Handler(BaseHTTPRequestHandler):

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path
        params = parse_qs(parsed.query)

        def p(key):
            v = params.get(key, [])
            return v[0] if v else ""

        if path == "/api/overview":
            self.json(analyzer.get_overview(p("start"), p("end")))
        elif path == "/api/tags/care-needs":
            self.json(tag_manager.get_care_needs())
        elif path == "/api/tags/untagged":
            self.json(tag_manager.get_untagged_customers(p("search")))
        elif path == "/api/tags/customers":
            self.json(tag_manager.get_tagged_customers())
        elif path == "/api/tags/auto/never-purchased":
            self.json(tag_manager.analyze_never_purchased())
        elif path.startswith("/api/tags/"):
            self.json({"error": "not found"}, 404)
        elif path == "/api/acquisition":
            self.json(analyzer.get_acquisition_detail(p("start"), p("end")))
        elif path == "/api/orders":
            self.json(analyzer.get_order_detail(p("start"), p("end")))
        elif path == "/api/config":
            self.json(analyzer.get_config_info())
        elif path == "/api/data/files":
            self.json(analyzer.list_data_files())
        elif path.startswith("/api/merge/download/"):
            session_id = path.split("/")[-1]
            self.download_merge(session_id)
        elif path.startswith("/api/dedup/download/"):
            session_id = path.split("/")[-1]
            self.download_dedup(session_id)
        elif path.startswith("/api/subtract/download/"):
            session_id = path.split("/")[-1]
            self.download_subtract(session_id)
        elif path == "/":
            self.serve("web/index.html")
        else:
            rel = path.lstrip("/")
            self.serve(rel)

    def do_POST(self):
        ct = self.headers.get("Content-Type", "")
        parsed = urlparse(self.path)

        if ct.startswith("multipart/form-data"):
            # 文件上传
            self.handle_multipart(parsed.path)
        else:
            # JSON body
            length = int(self.headers.get("Content-Length", 0))
            body = self.rfile.read(length).decode("utf-8") if length else "{}"
            try:
                data = json.loads(body)
            except json.JSONDecodeError:
                data = {}

            if parsed.path == "/api/tags/assign":
                tag_manager.assign_tag(data.get("customer_id"), data.get("tag"))
                self.json({"success": True})
            elif parsed.path == "/api/tags/remove":
                tag_manager.remove_tag(data.get("customer_id"), data.get("tag"))
                self.json({"success": True})
            elif parsed.path == "/api/report":
                md = analyzer.generate_report_md(
                    data.get("start", ""),
                    data.get("end", ""),
                    data.get("month", ""),
                )
                self.text(md, "text/markdown; charset=utf-8")
            else:
                self.json({"error": "not found"}, 404)

    # ── 文件上传 ──

    def handle_multipart(self, path):
        ct = self.headers.get("Content-Type", "")
        boundary = ct.split("boundary=")[-1].strip()
        length = int(self.headers.get("Content-Length", 0))
        raw = self.rfile.read(length)

        files = {}
        fields = {}

        for part in raw.split(("--" + boundary).encode()):
            if part.strip() in (b"", b"--"):
                continue
            idx = part.find(b"\r\n\r\n")
            if idx == -1:
                continue
            head = part[:idx].decode("utf-8", errors="replace")
            data = part[idx + 4:]
            if data.endswith(b"\r\n"):
                data = data[:-2]

            name = ""
            filename = ""
            for line in head.split("\r\n"):
                if "Content-Disposition" in line:
                    for seg in line.split(";"):
                        seg = seg.strip()
                        if seg.startswith('name="'):
                            name = seg[6:-1]
                        elif seg.startswith('filename="'):
                            filename = seg[10:-1]

            if filename:
                files[name] = {"filename": filename, "data": data}
            else:
                fields[name] = data.decode("utf-8", errors="replace")

        if path == "/api/merge/preview":
            self.handle_merge_preview(files)
        elif path == "/api/merge/execute":
            self.handle_merge_execute(files, fields)
        elif path == "/api/data/upload":
            self.handle_data_upload(files, fields)
        elif path == "/api/dedup/execute":
            self.handle_dedup_execute(files, fields)
        elif path == "/api/subtract/execute":
            self.handle_subtract_execute(files, fields)
        else:
            self.json({"error": "unknown path"}, 404)

    def handle_merge_preview(self, files):
        """上传一个文件并返回预览"""
        if "file" not in files:
            self.json({"error": "请上传文件"}, 400)
            return
        f = files["file"]
        tmp_path = TEMP / ("preview_" + f["filename"])
        tmp_path.write_bytes(f["data"])
        sheet_name = None  # 默认第一个
        try:
            info = analyzer.preview_table(str(tmp_path), sheet_name)
            info["temp_path"] = str(tmp_path)
            self.json(info)
        except Exception as e:
            self.json({"error": str(e)}, 400)

    def handle_merge_execute(self, files, fields):
        """上传两个文件 + 合并参数，执行合并"""
        for key in ["file_a", "file_b"]:
            if key not in files:
                self.json({"error": "请上传两个文件"}, 400)
                return

        fa = files["file_a"]
        fb = files["file_b"]
        path_a = TEMP / ("merge_a_" + fa["filename"])
        path_b = TEMP / ("merge_b_" + fb["filename"])
        path_a.write_bytes(fa["data"])
        path_b.write_bytes(fb["data"])

        key_a = fields.get("key_a", "")
        key_b = fields.get("key_b", "")
        join_type = fields.get("join_type", "inner")
        cols_a_raw = fields.get("cols_a", "")
        cols_b_raw = fields.get("cols_b", "")

        cols_a = json.loads(cols_a_raw) if cols_a_raw else None
        cols_b = json.loads(cols_b_raw) if cols_b_raw else None
        sheet_a = fields.get("sheet_a") or None
        sheet_b = fields.get("sheet_b") or None

        try:
            result = analyzer.merge_tables(
                str(path_a), str(path_b),
                key_a, key_b, join_type,
                cols_a, cols_b,
                sheet_a, sheet_b,
            )
            if "error" in result:
                self.json(result, 400)
                return

            # 缓存结果用于下载
            session_id = str(uuid.uuid4())[:8]
            _MERGE_CACHE[session_id] = {
                "headers": result["headers"],
                "rows": result["rows"],
            }

            # 取前 50 行预览
            result["preview"] = result["rows"][:50]
            result["session_id"] = session_id
            result["download_url"] = f"/api/merge/download/{session_id}"
            # 不返回全部数据（可能很大）
            del result["rows"]
            self.json(result)
        except Exception as e:
            self.json({"error": str(e)}, 400)

    def download_merge(self, session_id):
        """下载合并结果 CSV"""
        data = _MERGE_CACHE.get(session_id)
        if not data:
            self.send_error(404)
            return

        output = io.StringIO()
        writer = csv.writer(output)
        writer.writerow(data["headers"])
        for row in data["rows"]:
            writer.writerow(row)

        csv_bytes = output.getvalue().encode("utf-8-sig")
        self.send_response(200)
        self.send_header("Content-Type", "text/csv; charset=utf-8")
        self.send_header("Content-Length", str(len(csv_bytes)))
        self.send_header("Content-Disposition", f'attachment; filename="merge_result_{session_id}.csv"')
        self.end_headers()
        self.wfile.write(csv_bytes)

    # ── response helpers ──

    def json(self, data, code=200):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def text(self, text, mime="text/plain; charset=utf-8", code=200):
        body = text.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", mime)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def serve(self, rel):
        full = (BASE / rel).resolve()
        try:
            full.relative_to(BASE)
        except ValueError:
            self.send_error(403)
            return
        if not full.is_file():
            self.send_error(404)
            return
        ext = full.suffix.lower()
        mime = {
            ".html": "text/html; charset=utf-8",
            ".css": "text/css; charset=utf-8",
            ".js": "application/javascript; charset=utf-8",
        }.get(ext, "application/octet-stream")
        body = full.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", mime)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Cache-Control", "no-cache, no-store, must-revalidate")
        self.send_header("Pragma", "no-cache")
        self.send_header("Expires", "0")
        self.end_headers()
        self.wfile.write(body)

    def handle_data_upload(self, files, fields):
        """上传数据文件，自动替换系统数据源"""
        if "file" not in files:
            self.json({"error": "请上传文件"}, 400)
            return
        f = files["file"]
        fname = f["filename"]
        target_type = fields.get("type", "")  # "detail" or "order"

        if not fname.endswith(".xlsx"):
            self.json({"error": "仅支持 .xlsx 文件"}, 400)
            return

        # 保存到 data/ 目录
        data_dir = BASE / "data"
        data_dir.mkdir(parents=True, exist_ok=True)
        dest = data_dir / fname
        dest.write_bytes(f["data"])

        # 更新 config.json
        cfg_path = BASE / "config.json"
        cfg = {}
        if cfg_path.exists():
            cfg = json.loads(cfg_path.read_text(encoding="utf-8"))

        if target_type == "detail":
            cfg["detail_file"] = "system/data/" + fname
        elif target_type == "order":
            cfg["order_file"] = "system/data/" + fname
        else:
            # 自动识别
            if "明细" in fname or "带到店" in fname:
                cfg["detail_file"] = "system/data/" + fname
                target_type = "detail"
            elif "到店" in fname or "匹配活动" in fname:
                cfg["order_file"] = "system/data/" + fname
                target_type = "order"
            else:
                cfg["detail_file"] = "system/data/" + fname

        cfg_path.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")

        self.json({
            "success": True,
            "type": target_type,
            "filename": fname,
            "message": f"文件 {fname} 已上传并生效，请刷新页面查看最新数据",
        })

    def handle_dedup_execute(self, files, fields):
        if "file" not in files:
            self.json({"error": "请上传文件"}, 400)
            return
        f = files["file"]
        tmp_path = TEMP / ("dedup_" + f["filename"])
        tmp_path.write_bytes(f["data"])

        key_cols_raw = fields.get("key_columns", "[]")
        try:
            key_columns = json.loads(key_cols_raw)
        except json.JSONDecodeError:
            key_columns = []
        keep = fields.get("keep", "first")
        sheet_name = fields.get("sheet") or None

        if not key_columns:
            self.json({"error": "请选择要去重的列"}, 400)
            return

        try:
            result = analyzer.dedup_table(str(tmp_path), key_columns, sheet_name, keep)
            if "error" in result:
                self.json(result, 400)
                return

            session_id = str(uuid.uuid4())[:8]
            _DEDUP_CACHE[session_id] = {
                "headers": result["headers"],
                "rows": result["rows"],
            }
            result["preview"] = result["rows"][:50]
            result["session_id"] = session_id
            result["download_url"] = f"/api/dedup/download/{session_id}"
            del result["rows"]
            self.json(result)
        except Exception as e:
            self.json({"error": str(e)}, 400)

    def download_dedup(self, session_id):
        data = _DEDUP_CACHE.get(session_id)
        if not data:
            self.send_error(404)
            return
        import csv, io
        output = io.StringIO()
        writer = csv.writer(output)
        writer.writerow(data["headers"])
        for row in data["rows"]:
            writer.writerow(row)
        csv_bytes = output.getvalue().encode("utf-8-sig")
        self.send_response(200)
        self.send_header("Content-Type", "text/csv; charset=utf-8")
        self.send_header("Content-Length", str(len(csv_bytes)))
        self.send_header("Content-Disposition", f'attachment; filename="dedup_result_{session_id}.csv"')
        self.end_headers()
        self.wfile.write(csv_bytes)

    def handle_subtract_execute(self, files, fields):
        for k in ["master_file", "lookup_file"]:
            if k not in files:
                self.json({"error": f"请上传{k}文件"}, 400)
                return
        fa = files["master_file"]; fb = files["lookup_file"]
        pa = TEMP / ("sub_m_" + fa["filename"]); pb = TEMP / ("sub_l_" + fb["filename"])
        pa.write_bytes(fa["data"]); pb.write_bytes(fb["data"])
        km = fields.get("key_master", ""); kl = fields.get("key_lookup", "")
        if not km or not kl:
            self.json({"error": "请填写匹配列名"}, 400); return
        try:
            result = analyzer.subtract_table(str(pa), str(pb), km, kl)
            if "error" in result:
                self.json(result, 400); return
            sid = str(uuid.uuid4())[:8]
            _SUBTRACT_CACHE[sid] = {"headers": result["headers"], "rows": result["rows"]}
            result["preview"] = result["rows"][:50]; result["session_id"] = sid
            result["download_url"] = f"/api/subtract/download/{sid}"
            del result["rows"]; self.json(result)
        except Exception as e:
            self.json({"error": str(e)}, 400)

    def download_subtract(self, sid):
        data = _SUBTRACT_CACHE.get(sid)
        if not data:
            self.send_error(404); return
        import csv, io
        o = io.StringIO(); w = csv.writer(o)
        w.writerow(data["headers"])
        for row in data["rows"]:
            w.writerow(row)
        b = o.getvalue().encode("utf-8-sig")
        self.send_response(200)
        self.send_header("Content-Type", "text/csv; charset=utf-8")
        self.send_header("Content-Length", str(len(b)))
        self.send_header("Content-Disposition", f'attachment; filename="subtract_result_{sid}.csv"')
        self.end_headers(); self.wfile.write(b)

    def log_message(self, fmt, *args):
        print(f"  ⇨  {args[0]} {args[1]}")


def main():
    server = HTTPServer((HOST, PORT), Handler)
    print(f"╔════════════════════════════════════════╗")
    print(f"║  美田运营系统 v1.0                     ║")
    print(f"║                                        ║")
    print(f"║  启动地址: http://{HOST}:{PORT}          ║")
    print(f"║                                        ║")
    print(f"║  退出: Ctrl+C                          ║")
    print(f"╚════════════════════════════════════════╝")
    print()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n服务已停止")
        server.server_close()


if __name__ == "__main__":
    main()
