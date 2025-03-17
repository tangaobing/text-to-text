import pdfplumber
from docx import Document
from docx.shared import Inches
from pathlib import Path
import os
import asyncio
from pdf2image import convert_from_path
from PIL import Image
import tempfile
import logging
import traceback
import json
from .advanced_converter import precise_convert_pdf_to_word

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

async def convert_pdf_to_word(pdf_path: Path, output_path: Path, session_dir: Path, use_advanced=False, style_config=None) -> Path:
    """
    将PDF文件转换为Word文档
    
    Args:
        pdf_path: PDF文件路径
        output_path: 输出Word文件路径
        session_dir: 会话目录，用于存储临时文件
        use_advanced: 是否使用高级转换模式
        style_config: 样式配置（仅在高级模式下使用）
        
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
        # 使用精确布局转换器
        return await precise_convert_pdf_to_word(pdf_path, output_path, session_dir)
            
    except Exception as e:
        logger.error(f"转换过程中出错: {str(e)}")
        logger.error(traceback.format_exc())
        raise

async def extract_and_add_images(pdf_path: Path, page_num: int, doc: Document, images_dir: Path):
    """
    从PDF页面提取图片并添加到Word文档
    
    Args:
        pdf_path: PDF文件路径
        page_num: 页码（从1开始）
        doc: Word文档对象
        images_dir: 图片存储目录
    """
    logger.info(f"开始提取第 {page_num} 页的图片")
    
    # 使用pdf2image将PDF页面转换为图片
    loop = asyncio.get_event_loop()
    
    try:
        # 在线程池中执行CPU密集型操作
        images = await loop.run_in_executor(
            None, 
            lambda: convert_from_path(
                pdf_path, 
                first_page=page_num, 
                last_page=page_num,
                dpi=300
            )
        )
        
        if not images:
            logger.warning(f"第 {page_num} 页未提取到图片")
            return
        
        logger.info(f"第 {page_num} 页提取到 {len(images)} 张图片")
        
        # 保存并添加图片到文档
        for i, image in enumerate(images):
            img_path = images_dir / f"page{page_num}_{i}.png"
            
            # 在线程池中保存图片
            await loop.run_in_executor(None, lambda: image.save(img_path, "PNG"))
            logger.info(f"保存图片: {img_path.absolute()}")
            
            # 添加图片到Word文档
            doc.add_picture(str(img_path), width=Inches(6))  # 调整宽度以适应页面
            logger.info(f"图片已添加到Word文档")
    
    except Exception as e:
        logger.error(f"提取和添加图片时出错: {str(e)}")
        logger.error(traceback.format_exc())
        raise 