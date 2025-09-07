# app.py
import os
import time
import uuid
from flask import Flask, render_template, request, redirect, url_for, send_from_directory, flash
from werkzeug.utils import secure_filename

from analyzer import analyze_and_color_file2_complete

def ensure_dir(p):
    if not os.path.isdir(p):
        os.makedirs(p, exist_ok=True)

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
ensure_dir(UPLOAD_DIR)

ALLOWED_EXT = {'.xls', '.xlsx'}

def create_app():
    app = Flask(__name__)
    app.secret_key = os.environ.get("SECRET_KEY", "dev-secret")
    app.config['MAX_CONTENT_LENGTH'] = int(os.environ.get("MAX_CONTENT_LENGTH_MB", "200")) * 1024 * 1024  # 默认 200MB

    # 简单清理：每次访问首页清理 24h 之前的临时文件
    def cleanup_uploads(hours=24):
        now = time.time()
        threshold = now - hours * 3600
        for f in os.listdir(UPLOAD_DIR):
            try:
                fp = os.path.join(UPLOAD_DIR, f)
                if os.path.isfile(fp) and os.path.getmtime(fp) < threshold:
                    os.remove(fp)
            except Exception:
                pass

    @app.route("/", methods=["GET", "POST"])
    def upload():
        cleanup_uploads()

        if request.method == "POST":
            file1 = request.files.get("file1")
            file2 = request.files.get("file2")

            if not file1 or not file2:
                flash("请同时选择【上传分析依据文件】与【需要分析的文件】")
                return redirect(url_for("upload"))

            ext1 = os.path.splitext(file1.filename)[1].lower()
            ext2 = os.path.splitext(file2.filename)[1].lower()

            if ext1 not in ALLOWED_EXT or ext2 not in ALLOWED_EXT:
                flash("仅支持 .xls / .xlsx 文件")
                return redirect(url_for("upload"))

            # 保存上传
            token = uuid.uuid4().hex[:12]
            f1_name = secure_filename(f"{token}_file1{ext1}")
            f2_name = secure_filename(f"{token}_file2{ext2}")
            f1_path = os.path.join(UPLOAD_DIR, f1_name)
            f2_path = os.path.join(UPLOAD_DIR, f2_name)
            file1.save(f1_path)
            file2.save(f2_path)

            # 结果文件名
            out_name = secure_filename(f"{token}_分析结果.xlsx")
            out_path = os.path.join(UPLOAD_DIR, out_name)

            # 执行分析
            stats = {"yellow": 0, "orange": 0, "pending": 0}
            def progress_cb(pct, msg=""):
                # 可扩展为写入 Redis/文件后端，再在前端轮询显示进度
                pass

            try:
                ok, y, o, p = analyze_and_color_file2_complete(f1_path, f2_path, out_path, progress_callback=progress_cb)
                if not ok:
                    flash("分析失败")
                    return redirect(url_for("upload"))
                stats.update({"yellow": y, "orange": o, "pending": p})
            except Exception as e:
                flash(f"处理出错：{e}")
                return redirect(url_for("upload"))

            return render_template("result.html",
                                   token=token,
                                   output_filename=out_name,
                                   yellow=stats["yellow"],
                                   orange=stats["orange"],
                                   pending=stats["pending"])

        return render_template("upload.html")

    @app.route("/download/<path:filename>")
    def download(filename):
        return send_from_directory(UPLOAD_DIR, filename, as_attachment=True)

    return app

app = create_app()

if __name__ == "__main__":
    # 开发模式本地调试：python app.py
    port = int(os.environ.get("PORT", "9090"))
    app.run(host="0.0.0.0", port=port, debug=True)
