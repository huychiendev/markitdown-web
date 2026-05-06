import os
import uuid
import shutil
import json
import subprocess
import threading
from typing import List
from fastapi import FastAPI, File, UploadFile
from fastapi.responses import JSONResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware

app = FastAPI(title="MarkItDown Web UI")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

os.makedirs("static/conversions", exist_ok=True)
app.mount("/static", StaticFiles(directory="static"), name="static")


@app.on_event("startup")
def mark_stale_jobs_as_errors():
    """Detect jobs that were interrupted by a server crash (e.g. OOM) and mark them as errors."""
    try:
        conv_dir = "static/conversions"
        if not os.path.exists(conv_dir):
            return
        for entry in os.listdir(conv_dir):
            job_dir = os.path.join(conv_dir, entry)
            if os.path.isdir(job_dir):
                if not os.path.exists(os.path.join(job_dir, "success.txt")) and \
                   not os.path.exists(os.path.join(job_dir, "error.txt")) and \
                   not os.path.exists(os.path.join(job_dir, "success.json")):
                    with open(os.path.join(job_dir, "error.txt"), "w", encoding="utf-8") as f:
                        f.write("Server crashed during processing (likely OOM). Please try a smaller file.")
    except Exception as e:
        print(f"Error marking stale jobs: {e}")


def cleanup_old_jobs(directory="static/conversions", max_size_mb=100):
    try:
        job_dirs = []
        total_size = 0
        for entry in os.scandir(directory):
            if entry.is_dir():
                dir_size = sum(f.stat().st_size for f in os.scandir(entry.path) if f.is_file())
                job_dirs.append({"path": entry.path, "mtime": entry.stat().st_mtime, "size": dir_size})
                total_size += dir_size
        if total_size > max_size_mb * 1024 * 1024:
            job_dirs.sort(key=lambda x: x["mtime"])
            for job in job_dirs:
                shutil.rmtree(job["path"], ignore_errors=True)
                total_size -= job["size"]
                if total_size <= (max_size_mb * 0.8) * 1024 * 1024:
                    break
    except Exception as e:
        print(f"Cleanup error: {e}")


def _run_worker(job_dir, args):
    """Spawn worker.py in a subprocess. If it gets OOM-killed, write error.txt."""
    try:
        result = subprocess.run(
            ["python", "worker.py"] + args,
            timeout=300,  # 5 min hard limit
            capture_output=True, text=True
        )
        # If subprocess was killed by signal (e.g. OOM killer sends SIGKILL = -9)
        if result.returncode != 0:
            if not os.path.exists(os.path.join(job_dir, "success.txt")) and \
               not os.path.exists(os.path.join(job_dir, "error.txt")):
                error_msg = result.stderr.strip() if result.stderr else f"Worker exited with code {result.returncode}"
                with open(os.path.join(job_dir, "error.txt"), "w", encoding="utf-8") as f:
                    f.write(error_msg)
    except subprocess.TimeoutExpired:
        if not os.path.exists(os.path.join(job_dir, "error.txt")):
            with open(os.path.join(job_dir, "error.txt"), "w", encoding="utf-8") as f:
                f.write("Processing timed out (>5 minutes). File may be too large or complex.")
    except Exception as e:
        if not os.path.exists(os.path.join(job_dir, "error.txt")):
            with open(os.path.join(job_dir, "error.txt"), "w", encoding="utf-8") as f:
                f.write(str(e))


@app.head("/")
@app.get("/")
def serve_index():
    return FileResponse("static/index.html")


@app.post("/api/convert")
def convert_file(file: UploadFile = File(...)):
    cleanup_old_jobs()

    job_id = str(uuid.uuid4())
    job_dir = f"static/conversions/{job_id}"
    os.makedirs(job_dir, exist_ok=True)
    file_path = os.path.join(job_dir, file.filename)

    try:
        with open(file_path, "wb") as f:
            shutil.copyfileobj(file.file, f)

        # Spawn worker in a background thread (thread is cheap, the heavy work is in subprocess)
        t = threading.Thread(target=_run_worker, args=(job_dir, [job_id, job_dir, file_path, file.filename]))
        t.daemon = True
        t.start()

        return {"success": True, "status": "processing", "job_id": job_id}

    except Exception as e:
        return JSONResponse(status_code=500, content={"success": False, "error": str(e)})


@app.post("/api/convert_batch")
def convert_batch(files: List[UploadFile] = File(...)):
    cleanup_old_jobs()

    job_id = str(uuid.uuid4())
    job_dir = f"static/conversions/{job_id}"
    os.makedirs(job_dir, exist_ok=True)

    try:
        file_entries = []
        for file in files:
            file_path = os.path.join(job_dir, file.filename)
            with open(file_path, "wb") as f:
                shutil.copyfileobj(file.file, f)
            file_entries.append({"path": file_path, "name": file.filename})

        def _batch_worker():
            results_data = []
            for entry in file_entries:
                fp, fn = entry["path"], entry["name"]
                # Run each file in its own subprocess for maximum memory isolation
                try:
                    subprocess.run(
                        ["python", "worker.py", job_id, job_dir, fp, fn],
                        timeout=300, capture_output=True, text=True
                    )
                except Exception as e:
                    print(f"Error converting {fn}: {e}")

                # Read result
                md_filename = f"{os.path.splitext(fn)[0]}.md"
                md_path = os.path.join(job_dir, md_filename)
                if os.path.exists(md_path):
                    with open(md_path, "r", encoding="utf-8") as f:
                        results_data.append({"filename": fn, "markdown": f.read()})

            # Create ZIP
            zip_path = f"static/conversions/{job_id}_archive"
            shutil.make_archive(zip_path, 'zip', job_dir)

            # Write batch success
            with open(os.path.join(job_dir, "success.json"), "w", encoding="utf-8") as f:
                json.dump(results_data, f)

            # Clean up per-file success.txt markers (batch uses success.json)
            for txt in [f for f in os.listdir(job_dir) if f == "success.txt"]:
                os.remove(os.path.join(job_dir, txt))

        t = threading.Thread(target=_batch_worker)
        t.daemon = True
        t.start()

        return {"success": True, "status": "processing", "job_id": job_id}

    except Exception as e:
        return JSONResponse(status_code=500, content={"success": False, "error": str(e)})


@app.get("/api/status/{job_id}")
def check_status(job_id: str):
    job_dir = f"static/conversions/{job_id}"
    if not os.path.exists(job_dir):
        return {"success": False, "error": "Job not found"}

    if os.path.exists(os.path.join(job_dir, "error.txt")):
        with open(os.path.join(job_dir, "error.txt"), "r", encoding="utf-8") as f:
            return {"success": False, "error": f.read()}

    if os.path.exists(os.path.join(job_dir, "success.txt")):
        with open(os.path.join(job_dir, "success.txt"), "r") as f:
            md_filename = f.read().strip()
        with open(os.path.join(job_dir, md_filename), "r", encoding="utf-8") as f:
            markdown_text = f.read()
        return {
            "success": True, "status": "completed",
            "markdown": markdown_text,
            "download_url": f"/static/conversions/{job_id}/{md_filename}"
        }

    if os.path.exists(os.path.join(job_dir, "success.json")):
        with open(os.path.join(job_dir, "success.json"), "r", encoding="utf-8") as f:
            results_data = json.load(f)
        return {
            "success": True, "status": "completed",
            "results": results_data,
            "download_url": f"/static/conversions/{job_id}_archive.zip"
        }

    return {"success": True, "status": "processing"}


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app:app", host="0.0.0.0", port=8000, reload=True)
