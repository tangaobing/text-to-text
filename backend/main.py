"""
PDF转Word API - 简化版
只保留PDF转Word功能，使用pdf2docx实现
"""

from fastapi import FastAPI, UploadFile, File, HTTPException, BackgroundTasks, Form
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
import os
import uuid
from pathlib import Path
import time
from typing import Dict, Optional
import asyncio
import logging

from utils.pdf_converter import convert_pdf_to_word, cleanup_session_dir
from utils.file_manager import generate_filename, cleanup_old_files

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
MAX_FILE_SIZE = 100 * 1024 * 1024  # 100MB

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
    style_config: Optional[str] = Form(None),
    background_tasks: BackgroundTasks = None
):
    """
    上传PDF文件并开始转换为Word文档
    
    Args:
        file: 上传的PDF文件
        style_config: 样式配置（可选）
        background_tasks: 后台任务
        
    Returns:
        任务ID和状态信息
    """
    # 验证文件类型
    if not file.filename.lower().endswith('.pdf'):
        raise HTTPException(status_code=400, detail="只支持PDF文件")
    
    # 生成任务ID
    task_id = str(uuid.uuid4())
    
    # 创建会话目录
    session_dir = TEMP_DIR / task_id
    session_dir.mkdir(exist_ok=True)
    
    # 保存上传的文件
    input_path = session_dir / "input.pdf"
    with open(input_path, "wb") as f:
        content = await file.read()
        if len(content) > MAX_FILE_SIZE:
            raise HTTPException(status_code=400, detail=f"文件大小超过限制（最大{MAX_FILE_SIZE/1024/1024}MB）")
        f.write(content)
    
    # 生成输出文件名
    output_filename = generate_filename(file.filename)
    output_path = session_dir / output_filename
    
    # 创建任务状态
    conversion_tasks[task_id] = {
        "status": "pending",
        "progress": 0,
        "input_filename": file.filename,
        "output_filename": None,
        "error": None
    }
    
    # 启动后台转换任务
    background_tasks.add_task(process_conversion, task_id, input_path, output_path, session_dir, style_config)
    
    return {
        "task_id": task_id,
        "status": "pending",
        "message": "文件已上传，开始转换"
    }

async def process_conversion(task_id: str, input_path: Path, output_path: Path, session_dir: Path, style_config: Optional[str] = None):
    """
    处理PDF到Word的转换任务
    
    Args:
        task_id: 任务ID
        input_path: 输入PDF文件路径
        output_path: 输出Word文件路径
        session_dir: 会话目录
        style_config: 样式配置（可选）
    """
    try:
        # 更新任务状态为处理中
        conversion_tasks[task_id]["status"] = "processing"
        logger.info(f"开始处理任务 {task_id}")
        logger.info(f"输入文件路径: {input_path.absolute()}")
        logger.info(f"输出文件路径: {output_path.absolute()}")
        
        # 模拟进度更新
        for i in range(1, 10):
            conversion_tasks[task_id]["progress"] = i * 10
            logger.info(f"任务 {task_id} 进度: {i * 10}%")
            await asyncio.sleep(0.5)  # 模拟处理时间
        
        # 使用pdf2docx进行转换
        await convert_pdf_to_word(input_path, output_path, session_dir)
        
        # 检查文件是否生成
        if output_path.exists():
            logger.info(f"文件成功生成: {output_path.absolute()}, 大小: {output_path.stat().st_size} 字节")
        else:
            logger.error(f"文件未生成: {output_path.absolute()}")
            raise Exception("文件转换失败，未生成输出文件")
        
        # 更新任务状态为完成
        conversion_tasks[task_id]["status"] = "completed"
        conversion_tasks[task_id]["progress"] = 100
        conversion_tasks[task_id]["output_filename"] = output_path.name
        logger.info(f"任务 {task_id} 完成")
        
    except Exception as e:
        # 更新任务状态为失败
        logger.error(f"任务 {task_id} 失败: {str(e)}", exc_info=True)
        conversion_tasks[task_id]["status"] = "failed"
        conversion_tasks[task_id]["error"] = str(e)

@app.get("/status/{task_id}")
async def get_task_status(task_id: str):
    """获取转换任务的状态"""
    if task_id not in conversion_tasks:
        logger.warning(f"任务不存在: {task_id}")
        raise HTTPException(status_code=404, detail="任务不存在")
    
    return conversion_tasks[task_id]

@app.get("/download/{task_id}")
async def download_file(task_id: str, background_tasks: BackgroundTasks):
    """下载转换后的Word文档"""
    if task_id not in conversion_tasks:
        logger.warning(f"任务不存在: {task_id}")
        raise HTTPException(status_code=404, detail="任务不存在")
    
    task_info = conversion_tasks[task_id]
    
    if task_info["status"] != "completed":
        raise HTTPException(status_code=400, detail="任务尚未完成")
    
    if not task_info["output_filename"]:
        raise HTTPException(status_code=400, detail="输出文件不存在")
    
    output_path = TEMP_DIR / task_id / task_info["output_filename"]
    
    if not output_path.exists():
        raise HTTPException(status_code=404, detail="文件不存在")
    
    # 标记任务为已下载
    task_info["downloaded"] = True
    
    # 设置定时任务，延迟5秒后删除临时文件
    session_dir = TEMP_DIR / task_id
    background_tasks.add_task(cleanup_session_dir, str(session_dir), 5)
    
    return FileResponse(
        path=output_path,
        filename=task_info["output_filename"],
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

@app.get("/style-config")
async def get_style_config():
    """获取默认样式配置"""
    default_config = {
        "font_mappings": {
            "default": "微软雅黑",
            "serif": "宋体",
            "sans-serif": "微软雅黑",
            "monospace": "Consolas"
        },
        "preserve_images": True,
        "preserve_tables": True,
        "preserve_hyperlinks": True,
        "code_highlight": True,
        "math_support": True
    }
    return default_config

@app.get("/")
async def root():
    """API根路径，返回服务状态"""
    return {"message": "PDF转Word API服务正在运行"}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000) 