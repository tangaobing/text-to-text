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

def cleanup_session_dir(session_dir: Path) -> None:
    """
    清理单个会话目录
    
    Args:
        session_dir: 会话目录路径
    """
    if not session_dir.exists():
        logger.warning(f"会话目录不存在: {session_dir}")
        return
    
    try:
        # 不再需要关闭日志文件，因为日志文件现在存储在单独的日志目录中
        
        # 使用忽略错误的方式删除目录
        try:
            shutil.rmtree(session_dir, ignore_errors=True)
            logger.info(f"已删除会话目录: {session_dir}")
        except Exception as e:
            logger.warning(f"删除会话目录时出错: {str(e)}")
            
            # 如果删除失败，尝试逐个删除文件
            try:
                for root, dirs, files in os.walk(session_dir, topdown=False):
                    for file in files:
                        file_path = Path(root) / file
                        try:
                            file_path.unlink()
                            logger.info(f"已删除文件: {file_path}")
                        except Exception as file_e:
                            logger.warning(f"删除文件时出错: {file_path}, {str(file_e)}")
                    
                    for dir in dirs:
                        dir_path = Path(root) / dir
                        try:
                            dir_path.rmdir()
                            logger.info(f"已删除目录: {dir_path}")
                        except Exception as dir_e:
                            logger.warning(f"删除目录时出错: {dir_path}, {str(dir_e)}")
                
                # 最后尝试删除根目录
                try:
                    session_dir.rmdir()
                    logger.info(f"已删除根目录: {session_dir}")
                except Exception as root_e:
                    logger.warning(f"删除根目录时出错: {session_dir}, {str(root_e)}")
            except Exception as walk_e:
                logger.error(f"遍历删除文件时出错: {str(walk_e)}")
    except Exception as e:
        logger.error(f"清理会话目录时出错: {str(e)}")

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
                cleanup_session_dir(session_dir)
        except Exception as e:
            logger.error(f"清理目录 {session_dir} 时出错: {str(e)}")
    
    logger.info("临时文件清理完成") 