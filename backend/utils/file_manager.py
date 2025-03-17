"""
文件管理模块 - 简化版
只保留必要的文件管理功能
"""

import os
import random
import string
import time
from pathlib import Path
import logging
import shutil

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# 最大文件大小 (100MB)
MAX_FILE_SIZE = 100 * 1024 * 1024

def generate_filename(original_name: str) -> str:
    """
    生成唯一的输出文件名
    
    Args:
        original_name: 原始文件名
        
    Returns:
        生成的唯一文件名
    """
    # 提取基本文件名（不含扩展名）
    base_name = Path(original_name).stem
    
    # 生成随机后缀
    random_suffix = ''.join(random.choices(string.ascii_letters + string.digits, k=5))
    
    # 返回新文件名
    return f"{base_name}_{random_suffix}.docx"

def cleanup_old_files(temp_dir: Path, max_age_hours: int = 24) -> None:
    """
    清理旧的临时文件
    
    Args:
        temp_dir: 临时文件目录
        max_age_hours: 文件最大保留时间（小时）
    """
    if not temp_dir.exists():
        return
    
    current_time = time.time()
    max_age_seconds = max_age_hours * 3600
    
    logger.info(f"开始清理超过 {max_age_hours} 小时的临时文件")
    
    # 遍历临时目录中的所有子目录
    for session_dir in temp_dir.iterdir():
        if not session_dir.is_dir():
            continue
        
        try:
            # 检查目录的修改时间
            dir_mtime = session_dir.stat().st_mtime
            age_seconds = current_time - dir_mtime
            
            # 如果目录超过最大保留时间，则删除
            if age_seconds > max_age_seconds:
                logger.info(f"删除旧会话目录: {session_dir}")
                try:
                    shutil.rmtree(session_dir, ignore_errors=True)
                except Exception as e:
                    logger.error(f"删除目录 {session_dir} 时出错: {str(e)}")
        except Exception as e:
            logger.error(f"处理目录 {session_dir} 时出错: {str(e)}")
    
    logger.info("临时文件清理完成") 