#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文件类型检测和分类模块
提供精确的文件类型识别功能
"""

import os
import zipfile
import io
from typing import Tuple, Dict, Optional
import xml.etree.ElementTree as ET


class FileTypeDetector:
    """文件类型检测器"""
    
    def __init__(self):
        # 文件头签名映射
        self.file_signatures = {
            # Office文档 (新格式)
            b'\x50\x4B\x03\x04': 'zip_based',
            b'\x50\x4B\x05\x06': 'zip_based',
            b'\x50\x4B\x07\x08': 'zip_based',
            
            # Office文档 (旧格式)
            b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1': 'ole_compound',
            
            # PDF文档
            b'\x25\x50\x44\x46': 'pdf',
            
            # RTF文档
            b'{\\rtf1': 'rtf',
            b'{\\rtf': 'rtf',
            
            # 图像文件
            b'\xFF\xD8\xFF': 'jpeg',
            b'\x89\x50\x4E\x47\x0D\x0A\x1A\x0A': 'png',
            b'\x47\x49\x46\x38': 'gif',
            b'\x42\x4D': 'bmp',
            
            # 音频文件
            b'\x49\x44\x33': 'mp3',
            b'\xFF\xFB': 'mp3',
            b'\x52\x49\x46\x46': 'wav_or_avi',
            
            # 视频文件
            b'\x00\x00\x00\x18\x66\x74\x79\x70': 'mp4',
            b'\x00\x00\x00\x20\x66\x74\x79\x70': 'mp4',
            
            # 文本文件
            b'\xFF\xFE': 'utf16_le',
            b'\xFE\xFF': 'utf16_be',
            b'\xEF\xBB\xBF': 'utf8_bom',
            
            # 压缩文件
            b'\x1F\x8B': 'gzip',
            b'\x42\x5A\x68': 'bzip2',
            b'\x37\x7A\xBC\xAF\x27\x1C': '7zip',
            
            # 可执行文件
            b'\x4D\x5A': 'exe',
            b'\x7F\x45\x4C\x46': 'elf',
        }
        
        # MIME类型映射
        self.mime_types = {
            'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
            'doc': 'application/msword',
            'xls': 'application/vnd.ms-excel',
            'ppt': 'application/vnd.ms-powerpoint',
            'pdf': 'application/pdf',
            'rtf': 'application/rtf',
            'txt': 'text/plain',
            'jpeg': 'image/jpeg',
            'png': 'image/png',
            'gif': 'image/gif',
            'bmp': 'image/bmp',
            'mp3': 'audio/mpeg',
            'wav': 'audio/wav',
            'mp4': 'video/mp4',
            'avi': 'video/x-msvideo',
            'zip': 'application/zip',
            'exe': 'application/x-msdownload'
        }
        
        # 文件类型描述
        self.type_descriptions = {
            'docx': 'Microsoft Word文档',
            'xlsx': 'Microsoft Excel工作簿',
            'pptx': 'Microsoft PowerPoint演示文稿',
            'doc': 'Microsoft Word文档 (旧格式)',
            'xls': 'Microsoft Excel工作簿 (旧格式)',
            'ppt': 'Microsoft PowerPoint演示文稿 (旧格式)',
            'pdf': 'PDF文档',
            'rtf': 'RTF富文本文档',
            'txt': '文本文件',
            'jpeg': 'JPEG图像',
            'png': 'PNG图像',
            'gif': 'GIF图像',
            'bmp': 'BMP图像',
            'mp3': 'MP3音频',
            'wav': 'WAV音频',
            'mp4': 'MP4视频',
            'avi': 'AVI视频',
            'zip': 'ZIP压缩文件',
            'exe': '可执行文件',
            'unknown': '未知文件类型'
        }
    
    def detect_file_type(self, file_data: bytes, file_path: str = "") -> Tuple[str, str, str]:
        """
        检测文件类型
        
        Args:
            file_data: 文件二进制数据
            file_path: 文件路径（可选，用于辅助判断）
            
        Returns:
            (文件扩展名, 文件类型描述, MIME类型)
        """
        # 首先通过文件头签名检测
        detected_type = self._detect_by_signature(file_data)
        
        if detected_type:
            if detected_type == 'zip_based':
                # 进一步检测Office文档类型
                office_type = self._detect_office_document(file_data)
                if office_type:
                    ext = f'.{office_type}'
                    desc = self.type_descriptions.get(office_type, '未知Office文档')
                    mime = self.mime_types.get(office_type, 'application/octet-stream')
                    return ext, desc, mime
                else:
                    return '.zip', 'ZIP压缩文件', 'application/zip'
            
            elif detected_type == 'ole_compound':
                # 检测OLE复合文档类型
                ole_type = self._detect_ole_document(file_data)
                if ole_type:
                    ext = f'.{ole_type}'
                    desc = self.type_descriptions.get(ole_type, '未知OLE文档')
                    mime = self.mime_types.get(ole_type, 'application/octet-stream')
                    return ext, desc, mime
                else:
                    return '.ole', 'OLE复合文档', 'application/octet-stream'
            
            elif detected_type == 'wav_or_avi':
                # 进一步区分WAV和AVI
                if b'WAVE' in file_data[:100]:
                    return '.wav', 'WAV音频', 'audio/wav'
                elif b'AVI ' in file_data[:100]:
                    return '.avi', 'AVI视频', 'video/x-msvideo'
                else:
                    return '.riff', 'RIFF文件', 'application/octet-stream'
            
            else:
                # 直接映射的类型
                ext = f'.{detected_type}'
                desc = self.type_descriptions.get(detected_type, f'{detected_type.upper()}文件')
                mime = self.mime_types.get(detected_type, 'application/octet-stream')
                return ext, desc, mime
        
        # 通过文件路径扩展名检测
        if file_path:
            path_ext = self._detect_by_extension(file_path)
            if path_ext:
                desc = self.type_descriptions.get(path_ext, f'{path_ext.upper()}文件')
                mime = self.mime_types.get(path_ext, 'application/octet-stream')
                return f'.{path_ext}', desc, mime
        
        # 尝试文本检测
        if self._is_text_file(file_data):
            return '.txt', '文本文件', 'text/plain'
        
        # 默认为未知类型
        return '.bin', '未知文件类型', 'application/octet-stream'
    
    def _detect_by_signature(self, file_data: bytes) -> Optional[str]:
        """
        通过文件头签名检测文件类型
        """
        for signature, file_type in self.file_signatures.items():
            if file_data.startswith(signature):
                return file_type
        
        # 检查更长的签名
        if len(file_data) >= 8:
            # 检查MP4签名
            if file_data[4:8] == b'ftyp':
                return 'mp4'
        
        return None
    
    def _detect_office_document(self, file_data: bytes) -> Optional[str]:
        """
        检测Office文档类型（基于ZIP的新格式）
        """
        try:
            with zipfile.ZipFile(io.BytesIO(file_data), 'r') as zf:
                file_list = zf.namelist()
                
                # 检查特征文件
                if 'word/document.xml' in file_list:
                    return 'docx'
                elif 'xl/workbook.xml' in file_list:
                    return 'xlsx'
                elif 'ppt/presentation.xml' in file_list:
                    return 'pptx'
                elif '[Content_Types].xml' in file_list:
                    # 通过Content_Types.xml进一步判断
                    try:
                        content_types = zf.read('[Content_Types].xml')
                        if b'wordprocessingml' in content_types:
                            return 'docx'
                        elif b'spreadsheetml' in content_types:
                            return 'xlsx'
                        elif b'presentationml' in content_types:
                            return 'pptx'
                    except:
                        pass
                
                return None
        except:
            return None
    
    def _detect_ole_document(self, file_data: bytes) -> Optional[str]:
        """
        检测OLE复合文档类型（旧格式Office文档）
        """
        try:
            # 简单的OLE文档类型检测
            # 在文件的前几KB中查找特征字符串
            search_data = file_data[:4096].lower()
            
            if b'microsoft office word' in search_data or b'word.document' in search_data:
                return 'doc'
            elif b'microsoft office excel' in search_data or b'excel.sheet' in search_data:
                return 'xls'
            elif b'microsoft office powerpoint' in search_data or b'powerpoint.slide' in search_data:
                return 'ppt'
            elif b'wordpad' in search_data:
                return 'doc'
            
            return None
        except:
            return None
    
    def _detect_by_extension(self, file_path: str) -> Optional[str]:
        """
        通过文件扩展名检测文件类型
        """
        if not file_path:
            return None
        
        ext = os.path.splitext(file_path.lower())[1][1:]  # 去掉点号
        
        if ext in self.type_descriptions:
            return ext
        
        # 检查常见的扩展名变体
        ext_variants = {
            'htm': 'html',
            'jpeg': 'jpg',
            'tiff': 'tif',
            'mpeg': 'mpg'
        }
        
        return ext_variants.get(ext)
    
    def _is_text_file(self, file_data: bytes) -> bool:
        """
        检测是否为文本文件
        """
        try:
            # 尝试解码为UTF-8
            sample_size = min(1024, len(file_data))
            sample_data = file_data[:sample_size]
            
            # 检查是否包含过多的控制字符
            text = sample_data.decode('utf-8', errors='ignore')
            
            # 计算可打印字符的比例
            printable_chars = sum(1 for c in text if c.isprintable() or c.isspace())
            ratio = printable_chars / len(text) if text else 0
            
            return ratio > 0.7  # 如果70%以上是可打印字符，认为是文本文件
        except:
            return False
    
    def get_file_category(self, file_type: str) -> str:
        """
        获取文件类别
        
        Args:
            file_type: 文件类型（不含点号的扩展名）
            
        Returns:
            文件类别
        """
        categories = {
            'document': ['doc', 'docx', 'pdf', 'rtf', 'txt', 'odt'],
            'spreadsheet': ['xls', 'xlsx', 'ods', 'csv'],
            'presentation': ['ppt', 'pptx', 'odp'],
            'image': ['jpg', 'jpeg', 'png', 'gif', 'bmp', 'tiff', 'tif', 'svg'],
            'audio': ['mp3', 'wav', 'wma', 'aac', 'flac', 'ogg'],
            'video': ['mp4', 'avi', 'mkv', 'mov', 'wmv', 'flv', 'webm'],
            'archive': ['zip', 'rar', '7z', 'tar', 'gz', 'bz2'],
            'executable': ['exe', 'msi', 'deb', 'rpm', 'dmg']
        }
        
        for category, extensions in categories.items():
            if file_type.lower() in extensions:
                return category
        
        return 'other'
    
    def format_file_size(self, size_bytes: int) -> str:
        """
        格式化文件大小
        
        Args:
            size_bytes: 文件大小（字节）
            
        Returns:
            格式化的文件大小字符串
        """
        if size_bytes == 0:
            return "0 B"
        
        size_names = ["B", "KB", "MB", "GB", "TB"]
        import math
        i = int(math.floor(math.log(size_bytes, 1024)))
        p = math.pow(1024, i)
        s = round(size_bytes / p, 2)
        
        return f"{s} {size_names[i]}"