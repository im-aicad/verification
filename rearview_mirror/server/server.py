"""Rearview mirror regulation detection web server.

Provides a small FastAPI shell around
pycatia_regulation_reflection_point_detection.main so the CATIA workflow can be
started from a browser without changing the algorithm module.
"""

from __future__ import annotations

import asyncio
import json
import logging
import os
import re
import socket
import shutil
import subprocess
import sys
import tempfile
import traceback
from contextlib import asynccontextmanager
from datetime import datetime
from pathlib import Path
from typing import Any

from fastapi import FastAPI, File, Form, HTTPException, UploadFile, WebSocket, WebSocketDisconnect
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel


IS_FROZEN = bool(getattr(sys, "frozen", False))
RESOURCE_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
BASE_DIR = Path(sys.executable).resolve().parent if IS_FROZEN else Path(__file__).resolve().parent
PROJECT_DIR = BASE_DIR.parent
WEB_DIR = RESOURCE_DIR / "web" if IS_FROZEN else PROJECT_DIR / "web"
ALGORITHM_DIR = RESOURCE_DIR / "pycatia_regulation_reflection_point_detection" if IS_FROZEN else BASE_DIR / "pycatia_regulation_reflection_point_detection"
ALGORITHM_MAIN = ALGORITHM_DIR / "main.py"
ALGORITHM_WORKER = RESOURCE_DIR / "run_algorithm_worker.py" if IS_FROZEN else BASE_DIR / "run_algorithm_worker.py"
DEFAULT_CATPART = ALGORITHM_DIR / "Outside_Mirror_Regulation_Check.CATPart"
if not DEFAULT_CATPART.exists():
    DEFAULT_CATPART = ALGORITHM_DIR / "Outside_Mirror_Regulation_Check_0.CATPart"
UPLOADS_DIR = BASE_DIR / "uploads"
SESSION_FILE = BASE_DIR / "session.json"

for directory in (UPLOADS_DIR,):
    directory.mkdir(parents=True, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-7s  %(message)s",
)
logger = logging.getLogger("rearview_server")


def _detect_server_host() -> str:
    return os.environ.get("REARVIEW_SERVER_HOST", "0.0.0.0")


def _detect_server_port() -> int:
    value = os.environ.get("REARVIEW_SERVER_PORT", "8000")
    try:
        return int(value)
    except ValueError:
        return 8000


def _is_port_available(port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        return sock.connect_ex(("127.0.0.1", port)) != 0


def _select_server_port(preferred_port: int) -> int:
    if _is_port_available(preferred_port):
        return preferred_port

    if os.environ.get("REARVIEW_SERVER_PORT"):
        print(
            f"端口 {preferred_port} 已被占用。请关闭旧服务，或设置 "
            "REARVIEW_SERVER_PORT 使用其他端口。",
            file=sys.stderr,
        )
        raise SystemExit(1)

    for port in range(preferred_port + 1, preferred_port + 21):
        if _is_port_available(port):
            print(
                f"端口 {preferred_port} 已被占用，自动改用端口 {port}。"
                f"访问地址: http://localhost:{port}",
                file=sys.stderr,
            )
            return port

    print(
        f"端口 {preferred_port}-{preferred_port + 20} 均被占用，请关闭旧服务后重试。",
        file=sys.stderr,
    )
    raise SystemExit(1)


SERVER_HOST = _detect_server_host()
SERVER_PORT = _detect_server_port()
ACTIVE_SERVER_PORT = SERVER_PORT


class ConnectionManager:
    def __init__(self) -> None:
        self.active_connections: set[WebSocket] = set()

    async def connect(self, websocket: WebSocket) -> None:
        await websocket.accept()
        self.active_connections.add(websocket)

    def disconnect(self, websocket: WebSocket) -> None:
        self.active_connections.discard(websocket)

    async def broadcast(self, payload: dict[str, Any]) -> None:
        if not self.active_connections:
            return
        message = json.dumps(payload, ensure_ascii=False)
        stale: list[WebSocket] = []
        for websocket in self.active_connections:
            try:
                await websocket.send_text(message)
            except Exception:
                stale.append(websocket)
        for websocket in stale:
            self.disconnect(websocket)


manager = ConnectionManager()


class SessionState:
    def __init__(self) -> None:
        self.uploaded_file: Path | None = None
        self.uploaded_file_name: str | None = None
        self.running: bool = False
        self.completed: bool = False
        self.last_result: dict[str, Any] | None = None
        self.last_error: str | None = None
        self.saved_as_path: Path | None = None
        self.report_path: Path | None = None

    def to_dict(self) -> dict[str, Any]:
        return {
            "uploaded_file": str(self.uploaded_file) if self.uploaded_file else None,
            "uploaded_file_name": self.uploaded_file_name,
            "running": self.running,
            "completed": self.completed,
            "last_result": self.last_result,
            "last_error": self.last_error,
            "saved_as_path": str(self.saved_as_path) if self.saved_as_path else None,
            "report_path": str(self.report_path) if self.report_path else None,
        }

    def save(self) -> None:
        SESSION_FILE.write_text(
            json.dumps(self.to_dict(), ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    def load(self) -> None:
        if not SESSION_FILE.exists():
            return
        try:
            data = json.loads(SESSION_FILE.read_text(encoding="utf-8"))
            uploaded_file = data.get("uploaded_file")
            saved_as_path = data.get("saved_as_path")
            report_path = data.get("report_path")
            self.uploaded_file = Path(uploaded_file) if uploaded_file else None
            self.uploaded_file_name = data.get("uploaded_file_name")
            self.running = False
            self.completed = bool(data.get("completed"))
            self.last_result = data.get("last_result")
            self.last_error = data.get("last_error")
            self.saved_as_path = Path(saved_as_path) if saved_as_path else None
            self.report_path = Path(report_path) if report_path else None
        except Exception:
            logger.warning("session.json 读取失败，使用空状态")

    @property
    def active_file(self) -> Path:
        if self.uploaded_file and self.uploaded_file.exists():
            return self.uploaded_file
        return DEFAULT_CATPART


session = SessionState()
session.load()
current_task: asyncio.Task | None = None
catia_workflow_lock = asyncio.Lock()


def workflow_is_active() -> bool:
    return (
        (current_task is not None and not current_task.done())
        or catia_workflow_lock.locked()
    )


async def emit_log(message: str, level: str = "info") -> None:
    logger.info(message)
    await manager.broadcast(
        {
            "type": "log",
            "data": {
                "time": datetime.now().strftime("%H:%M:%S"),
                "level": level,
                "message": message,
            },
        }
    )


async def broadcast_captured_log(message: str, level: str = "info") -> None:
    await manager.broadcast(
        {
            "type": "log",
            "data": {
                "time": datetime.now().strftime("%H:%M:%S"),
                "level": level,
                "message": message,
            },
        }
    )


async def emit_status() -> None:
    await manager.broadcast({"type": "status", "data": get_status_payload()})


def get_status_payload() -> dict[str, Any]:
    if session.running and not workflow_is_active():
        session.running = False
        session.save()
    return {
        "running": session.running,
        "completed": session.completed,
        "uploaded_file_name": session.uploaded_file_name,
        "active_file": str(session.active_file),
        "saved_as_path": str(session.saved_as_path) if session.saved_as_path else None,
        "report_path": str(session.report_path) if session.report_path else None,
        "last_error": session.last_error,
        "has_result_file": bool(session.saved_as_path and session.saved_as_path.exists()),
        "has_report_file": bool(session.report_path and session.report_path.exists()),
    }


def sanitize_filename(filename: str) -> str:
    name = Path(filename).name
    name = re.sub(r"[^\w.\-\u4e00-\u9fff]+", "_", name, flags=re.UNICODE)
    return name or "upload.CATPart"


def get_catia_application(start_if_missing: bool = False) -> Any:
    try:
        import win32com.client  # type: ignore
    except Exception as exc:
        raise RuntimeError("当前环境未安装 pywin32，无法检测 CATIA") from exc

    try:
        return win32com.client.GetActiveObject("CATIA.Application")
    except Exception:
        if not start_if_missing:
            raise RuntimeError("CATIA 未打开")
    try:
        catia = win32com.client.Dispatch("CATIA.Application")
        catia.Visible = True
        return catia
    except Exception as exc:
        raise RuntimeError(f"启动 CATIA 失败: {exc}") from exc


def run_algorithm_sync(
    read_file_path: Path,
    loop: asyncio.AbstractEventLoop | None = None,
    feature_names: dict[str, str] | None = None,
) -> dict[str, Any]:
    if not IS_FROZEN and not ALGORITHM_WORKER.exists():
        raise FileNotFoundError(f"未找到检测子进程文件: {ALGORITHM_WORKER}")

    result_file = Path(tempfile.gettempdir()) / (
        f"rearview_result_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}.json"
    )
    config_file = Path(tempfile.gettempdir()) / (
        f"rearview_config_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}.json"
    )
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    env["REARVIEW_OUTPUT_DIR"] = str(BASE_DIR / "output")
    config_file.write_text(
        json.dumps(feature_names or {}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    if IS_FROZEN:
        command = [
            sys.executable,
            "--rearview-worker",
            str(read_file_path),
            str(result_file),
            str(config_file),
        ]
    else:
        command = [
            sys.executable,
            str(ALGORITHM_WORKER),
            str(read_file_path),
            str(result_file),
            str(config_file),
        ]
    process = subprocess.Popen(
        command,
        cwd=str(BASE_DIR),
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
        env=env,
    )

    def is_frontend_visible_worker_log(line: str) -> bool:
        """Only user-facing worker output is forwarded to the browser."""
        if line.startswith("[临时法规反射取点]"):
            return False
        if line.startswith("  ") and "到边界距离=" in line:
            return False
        return True

    output_tail: list[str] = []
    assert process.stdout is not None
    for raw_line in process.stdout:
        line = raw_line.rstrip("\r\n")
        if not line.strip():
            continue
        logger.info("[worker] %s", line)
        output_tail.append(line)
        output_tail = output_tail[-80:]
        if loop is not None and is_frontend_visible_worker_log(line):
            asyncio.run_coroutine_threadsafe(
                broadcast_captured_log(line),
                loop,
            )

    return_code = process.wait()
    try:
        config_file.unlink()
    except Exception:
        pass
    if result_file.exists():
        try:
            result = json.loads(result_file.read_text(encoding="utf-8"))
        finally:
            try:
                result_file.unlink()
            except Exception:
                pass
        if return_code != 0 and result.get("success", False):
            result["success"] = False
            result["error"] = f"检测子进程异常退出，退出码: {return_code}"
        return result

    detail = "\n".join(output_tail[-20:]).strip()
    error = f"检测子进程崩溃或未返回结果，退出码: {return_code}"
    if detail:
        error = f"{error}\n最后输出:\n{detail}"
    return {
        "success": False,
        "error": error,
        "worker_exit_code": return_code,
    }


async def run_detection_task(read_file_path: Path) -> None:
    async with catia_workflow_lock:
        session.running = True
        session.completed = False
        session.last_error = None
        session.saved_as_path = None
        session.report_path = None
        session.save()
        await emit_status()
        await emit_log(f"开始后视镜法规反射点检测: {read_file_path}")

        try:
            await emit_log("正在连接 CATIA 并执行几何构造，请保持 CATIA 可用")
            loop = asyncio.get_running_loop()
            result = await loop.run_in_executor(None, run_algorithm_sync, read_file_path, loop)
            session.last_result = result
            session.completed = bool(result.get("success"))
            session.saved_as_path = (
                Path(result["saved_as_path"]) if result.get("saved_as_path") else None
            )
            session.report_path = (
                Path(result["report_path"]) if result.get("report_path") else None
            )

            if result.get("success"):
                await emit_log("检测完成")
                if session.saved_as_path:
                    await emit_log(f"结果文件已另存: {session.saved_as_path}")
                if session.report_path:
                    await emit_log(f"报告文件已生成: {session.report_path}")
            else:
                session.last_error = str(result.get("error", "未知错误"))
                await emit_log(f"检测失败: {session.last_error}", level="error")
        except Exception as exc:
            session.last_error = str(exc)
            session.last_result = {
                "success": False,
                "error": str(exc),
                "traceback": traceback.format_exc(),
            }
            await emit_log(f"检测异常: {exc}", level="error")
        finally:
            session.running = False
            session.save()
            await emit_status()


@asynccontextmanager
async def lifespan(_app: FastAPI):
    logger.info("Rearview mirror server started")
    logger.info("访问地址: http://localhost:%s", ACTIVE_SERVER_PORT)
    yield


app = FastAPI(
    title="Rearview Mirror Regulation Detection",
    description="CATIA 后视镜法规反射点检测前后端服务",
    version="1.0.0",
    lifespan=lifespan,
)


@app.get("/", response_class=HTMLResponse)
async def index():
    index_file = WEB_DIR / "index.html"
    if not index_file.exists():
        raise HTTPException(status_code=404, detail="web/index.html 不存在")
    return HTMLResponse(
        index_file.read_text(encoding="utf-8"),
        headers={"Cache-Control": "no-store"},
    )


@app.get("/app.js")
async def app_js():
    return FileResponse(
        WEB_DIR / "app.js",
        media_type="application/javascript",
        headers={"Cache-Control": "no-store"},
    )


@app.get("/style.css")
async def style_css():
    return FileResponse(
        WEB_DIR / "style.css",
        media_type="text/css",
        headers={"Cache-Control": "no-store"},
    )


@app.post("/api/upload")
async def upload(file: UploadFile = File(...)):
    filename = sanitize_filename(file.filename or "")
    if Path(filename).suffix.casefold() != ".catpart":
        raise HTTPException(status_code=400, detail="请上传 .CATPart 文件")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    destination = UPLOADS_DIR / f"{timestamp}_{filename}"
    with destination.open("wb") as output:
        shutil.copyfileobj(file.file, output)

    session.uploaded_file = destination
    session.uploaded_file_name = filename
    session.running = False
    session.completed = False
    session.last_error = None
    session.last_result = None
    session.saved_as_path = None
    session.report_path = None
    session.save()
    await emit_log(f"已上传文件: {filename}")
    await emit_status()
    return {"ok": True, "filename": filename, "path": str(destination)}


@app.post("/verify")
async def verify_rearview(
    radar_file: UploadFile | None = File(None),
    bumper_file: UploadFile | None = File(None),
    ground_file: UploadFile | None = File(None),
    harness_file: UploadFile | None = File(None),
    input_parameter_geo_set_name: str | None = Form(None),
    left_mirror_feature_name: str | None = Form(None),
    right_mirror_feature_name: str | None = Form(None),
    left_eye_point_feature_name: str | None = Form(None),
    right_eye_point_feature_name: str | None = Form(None),
    ground_feature_name: str | None = Form(None),
    left_vehicle_width_line_feature_name: str | None = Form(None),
    right_vehicle_width_line_feature_name: str | None = Form(None),
):
    """Compatibility endpoint for the reference single-page frontend."""
    global current_task
    del bumper_file, ground_file, harness_file

    if current_task is not None and current_task.done():
        current_task = None
    if current_task is not None and not current_task.done():
        raise HTTPException(status_code=409, detail="检测流程正在运行，请稍后再试")
    if catia_workflow_lock.locked():
        raise HTTPException(status_code=503, detail="检测任务进行中，请稍后再试")

    if radar_file is None:
        raise HTTPException(status_code=400, detail="后视镜 CATPart 文件为必填项")

    filename = sanitize_filename(radar_file.filename or "")
    if Path(filename).suffix.casefold() != ".catpart":
        raise HTTPException(status_code=400, detail="请上传 .CATPart 文件")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    task_dir = UPLOADS_DIR / timestamp
    task_dir.mkdir(parents=True, exist_ok=True)
    read_file_path = task_dir / filename
    with read_file_path.open("wb") as output:
        shutil.copyfileobj(radar_file.file, output)
    feature_names = {
        "input_parameter_geo_set_name": input_parameter_geo_set_name,
        "left_mirror_feature_name": left_mirror_feature_name,
        "right_mirror_feature_name": right_mirror_feature_name,
        "left_eye_point_feature_name": left_eye_point_feature_name,
        "right_eye_point_feature_name": right_eye_point_feature_name,
        "ground_feature_name": ground_feature_name,
        "left_vehicle_width_line_feature_name": left_vehicle_width_line_feature_name,
        "right_vehicle_width_line_feature_name": right_vehicle_width_line_feature_name,
    }
    feature_names = {
        key: value.strip()
        for key, value in feature_names.items()
        if value is not None and value.strip()
    }

    async with catia_workflow_lock:
        current_task = asyncio.current_task()
        session.uploaded_file = read_file_path
        session.uploaded_file_name = filename
        session.running = True
        session.completed = False
        session.last_error = None
        session.last_result = None
        session.saved_as_path = None
        session.report_path = None
        session.save()
        await emit_status()
        await emit_log(f"开始后视镜法规反射点检测: {read_file_path}")

        try:
            loop = asyncio.get_running_loop()
            result = await loop.run_in_executor(
                None,
                run_algorithm_sync,
                read_file_path,
                loop,
                feature_names,
            )
            session.last_result = result
            session.completed = bool(result.get("success"))
            session.saved_as_path = (
                Path(result["saved_as_path"]) if result.get("saved_as_path") else None
            )
            session.report_path = (
                Path(result["report_path"]) if result.get("report_path") else None
            )
            if not result.get("success"):
                session.last_error = str(result.get("error", "未知错误"))
                await emit_log(f"检测失败: {session.last_error}", level="error")
                raise HTTPException(status_code=500, detail=session.last_error)
            if not session.saved_as_path or not session.saved_as_path.exists():
                raise HTTPException(status_code=500, detail="检测完成但未生成结果文件")
            if not session.report_path or not session.report_path.exists():
                raise HTTPException(status_code=500, detail="检测完成但未生成报告文件")

            await emit_log("检测完成")
            await emit_log(f"结果文件已另存: {session.saved_as_path}")
            await emit_log(f"报告文件已生成: {session.report_path}")
            return {
                "ok": True,
                "result_filename": session.saved_as_path.name,
                "report_filename": session.report_path.name,
                "result_download_url": "/api/download-result",
                "report_download_url": "/api/download-report",
            }
        except HTTPException:
            raise
        except Exception as exc:
            session.last_error = str(exc)
            session.last_result = {
                "success": False,
                "error": str(exc),
                "traceback": traceback.format_exc(),
            }
            await emit_log(f"检测异常: {exc}", level="error")
            raise HTTPException(
                status_code=500,
                detail=f"检测失败: {exc}\n{traceback.format_exc()}",
            ) from exc
        finally:
            session.running = False
            session.save()
            current_task = None
            await emit_status()


class StartRequest(BaseModel):
    use_default_file: bool = False


@app.post("/api/start")
async def start(req: StartRequest | None = None):
    global current_task
    if current_task is not None and current_task.done():
        current_task = None
    if current_task is not None and not current_task.done():
        raise HTTPException(status_code=409, detail="检测流程正在运行")

    read_file_path = DEFAULT_CATPART if req and req.use_default_file else session.active_file
    if not read_file_path.exists():
        raise HTTPException(status_code=400, detail=f"CATPart 文件不存在: {read_file_path}")

    current_task = asyncio.create_task(run_detection_task(read_file_path))
    return {"ok": True, "message": "检测流程已启动", "active_file": str(read_file_path)}


@app.post("/api/reset")
async def reset():
    global current_task
    if current_task is not None and current_task.done():
        current_task = None
    if current_task is not None and not current_task.done():
        raise HTTPException(status_code=409, detail="检测流程正在运行，暂不能重置")
    if catia_workflow_lock.locked():
        raise HTTPException(status_code=409, detail="检测流程正在运行，暂不能重置")
    session.uploaded_file = None
    session.uploaded_file_name = None
    session.running = False
    session.completed = False
    session.last_error = None
    session.last_result = None
    session.saved_as_path = None
    session.report_path = None
    session.save()
    await emit_log("已重置，请重新上传后视镜法规反射点检测文件")
    await emit_status()
    return {"ok": True}


@app.get("/api/status")
async def status():
    return get_status_payload()


@app.get("/api/catia-status")
async def catia_status():
    try:
        catia = get_catia_application(start_if_missing=False)
        return {
            "running": True,
            "visible": bool(getattr(catia, "Visible", False)),
            "message": "CATIA 已打开",
        }
    except Exception as exc:
        return {"running": False, "visible": False, "message": str(exc)}


@app.post("/api/open-catia")
async def open_catia():
    try:
        catia = get_catia_application(start_if_missing=True)
        try:
            catia.Visible = True
        except Exception:
            pass
        await emit_log("CATIA 已打开或已连接", level="success")
        return {
            "ok": True,
            "running": True,
            "visible": bool(getattr(catia, "Visible", False)),
            "message": "CATIA 已打开或已连接",
        }
    except Exception as exc:
        await emit_log(f"CATIA 打开失败: {exc}", level="error")
        raise HTTPException(status_code=500, detail=str(exc)) from exc


@app.get("/health")
async def health():
    return {"status": "ok"}


@app.get("/api/result")
async def result():
    if session.last_result is None:
        raise HTTPException(status_code=404, detail="暂无检测结果")
    return session.last_result


@app.get("/api/download-result")
async def download_result():
    if not session.saved_as_path or not session.saved_as_path.exists():
        raise HTTPException(status_code=404, detail="暂无可下载的结果文件")
    return FileResponse(
        session.saved_as_path,
        media_type="application/octet-stream",
        filename=session.saved_as_path.name,
    )


@app.get("/api/download-report")
async def download_report():
    if not session.report_path or not session.report_path.exists():
        raise HTTPException(status_code=404, detail="暂无可下载的报告文件")
    return FileResponse(
        session.report_path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=session.report_path.name,
    )


@app.websocket("/ws")
async def websocket_endpoint(websocket: WebSocket):
    await manager.connect(websocket)
    await websocket.send_text(
        json.dumps({"type": "status", "data": get_status_payload()}, ensure_ascii=False)
    )
    try:
        while True:
            await websocket.receive_text()
    except WebSocketDisconnect:
        manager.disconnect(websocket)
    except Exception:
        manager.disconnect(websocket)


if WEB_DIR.exists():
    lib_dir = WEB_DIR / "lib"
    if lib_dir.exists():
        app.mount("/lib", StaticFiles(directory=str(lib_dir)), name="lib")


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--rearview-worker":
        from run_algorithm_worker import main as worker_main

        sys.argv = [sys.argv[0], *sys.argv[2:]]
        raise SystemExit(worker_main())

    import uvicorn

    selected_port = _select_server_port(SERVER_PORT)
    ACTIVE_SERVER_PORT = selected_port

    uvicorn.run(
        app,
        host=SERVER_HOST,
        port=selected_port,
        reload=False,
        log_level="info",
    )
