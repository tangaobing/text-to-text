"""
PDF到Word转换模块 - 使用pdf2docx实现
简化版本，只保留核心功能
"""

import logging
from pathlib import Path
import time
import traceback
import asyncio
import shutil
import os

# 导入pdf2docx模块
from pdf2docx import Converter

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

async def convert_pdf_to_word(pdf_path: Path, output_path: Path, session_dir: Path) -> Path:
    """
    将PDF文件转换为Word文档
    
    Args:
        pdf_path: PDF文件路径
        output_path: 输出Word文件路径
        session_dir: 会话目录，用于存储临时文件
        
    Returns:
        输出Word文件的路径
    """
    logger.info(f"开始转换: {pdf_path} -> {output_path}")
    
    # 检查输入文件是否存在
    if not pdf_path.exists():
        error_msg = f"输入PDF文件不存在: {pdf_path}"
        logger.error(error_msg)
        raise FileNotFoundError(error_msg)
    
    # 检查文件大小
    file_size = pdf_path.stat().st_size
    logger.info(f"PDF文件大小: {file_size} 字节")
    if file_size == 0:
        error_msg = "PDF文件为空"
        logger.error(error_msg)
        raise ValueError(error_msg)
    
    try:
        # 创建日志目录
        logs_dir = Path("logs")
        logs_dir.mkdir(exist_ok=True)
        
        # 记录转换开始时间
        start_time = time.time()
        
        # 使用pdf2docx进行转换
        cv = Converter(str(pdf_path))
        cv.convert(
            str(output_path), 
            start=0, 
            end=None,
            multi_paragraphs=True,  # 启用多段落支持
            debug=False,
            min_paragraph_height=5.0,  # 降低段落高度阈值
            connected_border_tolerance=3.0,  # 提高边框连接容差
            line_overlap_threshold=0.9,  # 提高行重叠阈值
            line_break_width_ratio=0.5,  # 调整行断开宽度比例
            float_image_ignorable_gap=5.0,  # 减小浮动图像可忽略间隙
            page_margin_factor=0.0,  # 减小页边距因子
            detect_table_bottom_border=False,  # 禁用表格底部边框检测，防止文本框被识别为表格
            text_box_width_limit=20.0,  # 增加文本框宽度限制，更好地识别文本框
            text_box_free_text_tolerance=0.5,  # 降低自由文本容差，更好地处理文本框内容
            extract_stream_table=False  # 禁用流式表格提取，防止文本框被识别为表格
        )
        cv.close()
        
        # 记录转换结束时间
        end_time = time.time()
        logger.info(f"转换完成，耗时: {end_time - start_time:.2f}秒")
        
        # 验证输出文件
        if output_path.exists() and output_path.stat().st_size > 0:
            logger.info(f"成功生成Word文档: {output_path}")
            return output_path
        else:
            error_msg = f"转换失败: 未生成有效的Word文档"
            logger.error(error_msg)
            raise RuntimeError(error_msg)
            
    except Exception as e:
        logger.error(f"转换过程中出错: {str(e)}")
        logger.error(traceback.format_exc())
        raise

async def cleanup_session_dir(session_dir: str, delay_seconds: int = 5):
    """
    延迟指定秒数后删除会话目录
    
    Args:
        session_dir: 会话目录路径
        delay_seconds: 延迟秒数
    """
    try:
        logger.info(f"计划在 {delay_seconds} 秒后清理临时文件: {session_dir}")
        await asyncio.sleep(delay_seconds)
        
        # 删除会话目录及其所有内容
        session_dir_path = Path(session_dir)
        if session_dir_path.exists():
            try:
                shutil.rmtree(session_dir_path, ignore_errors=True)
                logger.info(f"已删除会话目录: {session_dir}")
            except Exception as e:
                logger.warning(f"删除会话目录时出错: {str(e)}")
                # 如果删除失败，尝试逐个删除文件
                try:
                    for root, dirs, files in os.walk(str(session_dir_path), topdown=False):
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
                except Exception as walk_e:
                    logger.error(f"遍历删除文件时出错: {str(walk_e)}")
    except Exception as e:
        logger.error(f"清理临时文件时出错: {str(e)}") 