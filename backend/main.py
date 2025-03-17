from fastapi import FastAPI, UploadFile, File, HTTPException, BackgroundTasks, Request, Form, Body
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
import os
import uuid
import shutil
from pathlib import Path
import time
from typing import Dict, Optional, List, Any
import asyncio
import logging
import json

from utils.converter import convert_pdf_to_word
from utils.markdown_converter import convert_markdown_to_word
from utils.file_manager import generate_filename, cleanup_old_files, cleanup_session_dir

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

app = FastAPI(title="PDF转Word API", description="将PDF文件转换为Word文档的API服务")

# 配置CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # 在生产环境中应该限制为前端域名
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 创建临时目录
TEMP_DIR = Path("temp")
TEMP_DIR.mkdir(exist_ok=True)
logger.info(f"临时目录路径: {TEMP_DIR.absolute()}")

# 创建日志目录
LOGS_DIR = Path("logs")
LOGS_DIR.mkdir(exist_ok=True)
logger.info(f"日志目录路径: {LOGS_DIR.absolute()}")

# 存储转换任务状态
conversion_tasks: Dict[str, Dict] = {}

# 配置
MAX_FILE_SIZE = 20 * 1024 * 1024  # 20MB

# 定期清理临时文件
@app.on_event("startup")
async def startup_event():
    # 启动时清理旧文件
    cleanup_old_files(TEMP_DIR, max_age_hours=24)
    
    # 启动定期清理任务
    asyncio.create_task(periodic_cleanup())

async def periodic_cleanup():
    """定期清理临时文件和日志文件"""
    while True:
        try:
            # 每小时清理一次
            await asyncio.sleep(3600)
            logger.info("开始定期清理临时文件和日志文件")
            
            # 清理临时文件
            cleanup_old_files(TEMP_DIR, max_age_hours=24)
            
            # 清理日志文件
            try:
                current_time = time.time()
                max_age_seconds = 24 * 3600  # 24小时
                
                # 遍历日志目录中的所有文件
                for log_file in LOGS_DIR.glob("*.log"):
                    try:
                        # 检查文件的修改时间
                        file_mtime = log_file.stat().st_mtime
                        age_seconds = current_time - file_mtime
                        
                        # 如果文件超过最大保留时间，则删除
                        if age_seconds > max_age_seconds:
                            logger.info(f"删除旧日志文件: {log_file}")
                            log_file.unlink()
                    except Exception as e:
                        logger.error(f"清理日志文件 {log_file} 时出错: {str(e)}")
                
                logger.info("日志文件清理完成")
            except Exception as e:
                logger.error(f"清理日志文件时出错: {str(e)}")
        except Exception as e:
            logger.error(f"定期清理任务时出错: {str(e)}")

@app.post("/upload")
async def upload_file(
    file: UploadFile = File(...), 
    conversion_type: str = Form("pdf-to-word"),  # 新增参数：转换类型
    style_config: Optional[str] = Form(None),
    background_tasks: BackgroundTasks = None
):
    """
    上传文件并开始转换
    
    Args:
        file: 上传的文件（PDF或Markdown）
        conversion_type: 转换类型（pdf-to-word, markdown-to-word等）
        style_config: 样式配置JSON字符串（可选）
    """
    logger.info(f"收到文件上传请求: {file.filename}, 转换类型: {conversion_type}")
    
    # 验证文件大小
    file_size = 0
    chunk_size = 1024 * 1024  # 1MB chunks
    while True:
        chunk = await file.read(chunk_size)
        if not chunk:
            break
        file_size += len(chunk)
        if file_size > MAX_FILE_SIZE:
            raise HTTPException(status_code=413, detail=f"文件大小超过限制（最大20MB）")
    await file.seek(0)  # 重置文件指针到开始位置
    
    # 验证文件类型和转换类型的匹配
    file_extension = Path(file.filename).suffix.lower()
    
    if conversion_type == "pdf-to-word" and file_extension != ".pdf":
        logger.warning(f"文件类型不匹配: {file.filename} 不是PDF文件")
        raise HTTPException(status_code=400, detail="PDF转Word需要上传PDF文件")
    
    if conversion_type == "markdown-to-word" and file_extension not in [".md", ".markdown"]:
        logger.warning(f"文件类型不匹配: {file.filename} 不是Markdown文件")
        raise HTTPException(status_code=400, detail="Markdown转Word需要上传Markdown文件")
    
    # 创建唯一的会话ID
    session_id = str(uuid.uuid4())
    session_dir = TEMP_DIR / session_id
    session_dir.mkdir(exist_ok=True)
    logger.info(f"创建会话目录: {session_dir.absolute()}")
    
    # 保存上传的文件
    input_filename = "input" + file_extension
    input_path = session_dir / input_filename
    with open(input_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)
    logger.info(f"保存上传文件: {input_path.absolute()}")
    
    # 解析样式配置（如果提供）
    parsed_style_config = None
    if style_config:
        try:
            parsed_style_config = json.loads(style_config)
            logger.info(f"解析样式配置: {parsed_style_config}")
        except json.JSONDecodeError as e:
            logger.warning(f"样式配置解析失败: {str(e)}")
            # 继续处理，让转换器自动检测样式
    
    # 创建转换任务
    task_id = str(uuid.uuid4())
    conversion_tasks[task_id] = {
        "status": "processing",
        "progress": 0,
        "session_id": session_id,
        "original_filename": file.filename,
        "created_at": time.time(),
        "downloaded": False,  # 添加下载状态标记
        "style_config": parsed_style_config,  # 记录样式配置
        "conversion_type": conversion_type  # 记录转换类型
    }
    logger.info(f"创建转换任务: {task_id}")
    
    # 在后台执行转换
    background_tasks.add_task(
        process_conversion, 
        task_id=task_id,
        input_path=input_path,
        session_dir=session_dir,
        conversion_type=conversion_type,
        style_config=parsed_style_config
    )
    
    return {"task_id": task_id}

async def process_conversion(
    task_id: str, 
    input_path: Path, 
    session_dir: Path,
    conversion_type: str,
    style_config: Optional[Dict[str, Any]] = None
):
    """后台处理文件转换任务"""
    try:
        # 更新任务状态
        conversion_tasks[task_id]["status"] = "processing"
        logger.info(f"开始处理转换任务: {task_id}, 类型: {conversion_type}")
        
        # 执行转换
        output_filename = generate_filename(conversion_tasks[task_id]["original_filename"])
        output_path = session_dir / output_filename
        logger.info(f"输出文件路径: {output_path.absolute()}")
        
        # 模拟进度更新
        for i in range(1, 10):
            conversion_tasks[task_id]["progress"] = i * 10
            logger.info(f"任务 {task_id} 进度: {i * 10}%")
            await asyncio.sleep(0.5)  # 模拟处理时间
        
        # 根据转换类型选择不同的转换函数
        if conversion_type == "pdf-to-word":
            logger.info(f"使用PDF转Word转换器")
            await convert_pdf_to_word(
                input_path, 
                output_path, 
                session_dir,
                style_config=style_config
            )
        elif conversion_type == "markdown-to-word":
            logger.info(f"使用Markdown转Word转换器")
            await convert_markdown_to_word(
                input_path, 
                output_path, 
                session_dir,
                style_config=style_config
            )
        else:
            logger.error(f"不支持的转换类型: {conversion_type}")
            raise ValueError(f"不支持的转换类型: {conversion_type}")
        
        # 检查文件是否生成
        if output_path.exists():
            logger.info(f"文件成功生成: {output_path.absolute()}, 大小: {output_path.stat().st_size} 字节")
        else:
            logger.error(f"文件未生成: {output_path.absolute()}")
            raise Exception("文件转换失败，未生成输出文件")
        
        # 更新任务状态为完成
        conversion_tasks[task_id]["status"] = "completed"
        conversion_tasks[task_id]["progress"] = 100
        conversion_tasks[task_id]["output_filename"] = output_filename
        logger.info(f"任务 {task_id} 完成")
        
    except Exception as e:
        # 更新任务状态为失败
        logger.error(f"任务 {task_id} 失败: {str(e)}", exc_info=True)
        conversion_tasks[task_id]["status"] = "failed"
        conversion_tasks[task_id]["error"] = str(e)

@app.get("/status/{task_id}")
async def get_task_status(task_id: str):
    """获取转换任务的状态"""
    logger.info(f"获取任务状态: {task_id}")
    if task_id not in conversion_tasks:
        logger.warning(f"任务不存在: {task_id}")
        raise HTTPException(status_code=404, detail="任务不存在")
    
    return conversion_tasks[task_id]

@app.get("/download/{task_id}")
async def download_file(task_id: str, background_tasks: BackgroundTasks):
    """
    下载转换后的Word文档
    
    Args:
        task_id: 转换任务ID
    """
    logger.info(f"收到下载请求: {task_id}")
    
    if task_id not in conversion_tasks:
        logger.warning(f"任务不存在: {task_id}")
        raise HTTPException(status_code=404, detail="任务不存在")
    
    task = conversion_tasks[task_id]
    
    if task["status"] != "completed":
        logger.warning(f"任务未完成: {task_id}, 状态: {task['status']}")
        raise HTTPException(status_code=400, detail="任务尚未完成")
    
    session_dir = TEMP_DIR / task["session_id"]
    output_filename = task["output_filename"]
    file_path = session_dir / output_filename
    
    if not file_path.exists():
        logger.error(f"文件不存在: {file_path.absolute()}")
        raise HTTPException(status_code=404, detail="文件不存在")
    
    # 标记任务为已下载
    conversion_tasks[task_id]["downloaded"] = True
    logger.info(f"标记任务为已下载: {task_id}")
    
    # 添加后台任务，延迟3秒后清理临时文件
    background_tasks.add_task(delayed_cleanup, task_id, session_dir, delay=3)
    
    return FileResponse(
        path=file_path, 
        filename=output_filename,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

async def delayed_cleanup(task_id: str, session_dir: Path, delay: int = 3):
    """延迟指定秒数后清理临时文件"""
    logger.info(f"计划在{delay}秒后清理任务: {task_id}")
    await asyncio.sleep(delay)
    
    try:
        # 清理会话目录
        cleanup_session_dir(session_dir)
        logger.info(f"已自动清理会话目录: {session_dir}")
        
        # 从任务列表中移除任务
        if task_id in conversion_tasks:
            del conversion_tasks[task_id]
            logger.info(f"已从任务列表中移除任务: {task_id}")
    except Exception as e:
        logger.error(f"自动清理时出错: {str(e)}")

# 简化样式配置接口
@app.get("/style-config")
async def get_style_config():
    """获取默认样式配置"""
    # 返回默认配置
    return {
        "font_mapping": {
            "Helvetica": "微软雅黑",
            "TimesNewRoman": "宋体",
            "Times-Roman": "宋体",
            "Times": "宋体",
            "Arial": "微软雅黑",
            "Courier": "仿宋",
            "Symbol": "Symbol"
        },
        "color_threshold": 0.9,
        "line_spacing": 1.5,
        "paragraph_indent": "2字符"
    }

# 添加一个简单的API根路径响应
@app.get("/")
async def root():
    return {"message": "PDF转Word API服务正在运行"}

# 挂载静态文件目录（用于前端构建文件）
# 注释掉这一行，因为前端dist目录可能还不存在
# app.mount("/", StaticFiles(directory="../frontend/dist", html=True), name="static")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True) 