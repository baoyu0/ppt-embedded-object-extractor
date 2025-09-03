#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
异常处理和错误报告模块
提供统一的错误处理和日志记录功能
"""

import os
import sys
import traceback
import logging
from datetime import datetime
from typing import Dict, List, Optional, Any
from enum import Enum


class ErrorLevel(Enum):
    """错误级别枚举"""
    INFO = "INFO"
    WARNING = "WARNING"
    ERROR = "ERROR"
    CRITICAL = "CRITICAL"


class ErrorCode(Enum):
    """错误代码枚举"""
    # 文件相关错误
    FILE_NOT_FOUND = "E001"
    FILE_ACCESS_DENIED = "E002"
    FILE_CORRUPTED = "E003"
    FILE_FORMAT_UNSUPPORTED = "E004"
    
    # 目录相关错误
    DIR_NOT_FOUND = "E101"
    DIR_ACCESS_DENIED = "E102"
    DIR_CREATE_FAILED = "E103"
    
    # 提取相关错误
    EXTRACTION_FAILED = "E201"
    NO_EMBEDDED_OBJECTS = "E202"
    OBJECT_CORRUPTED = "E203"
    SAVE_FAILED = "E204"
    
    # 系统相关错误
    MEMORY_ERROR = "E301"
    DISK_SPACE_ERROR = "E302"
    PERMISSION_ERROR = "E303"
    
    # 其他错误
    UNKNOWN_ERROR = "E999"


class PPTExtractorError(Exception):
    """PPT提取器自定义异常类"""
    
    def __init__(self, message: str, error_code: ErrorCode = ErrorCode.UNKNOWN_ERROR, 
                 details: Optional[Dict[str, Any]] = None):
        self.message = message
        self.error_code = error_code
        self.details = details or {}
        self.timestamp = datetime.now()
        super().__init__(self.message)
    
    def __str__(self):
        return f"[{self.error_code.value}] {self.message}"


class ErrorHandler:
    """错误处理器"""
    
    def __init__(self, log_file: Optional[str] = None, enable_console: bool = True):
        self.errors = []
        self.warnings = []
        self.log_file = log_file
        self.enable_console = enable_console
        
        # 设置日志记录器
        self._setup_logger()
    
    def _setup_logger(self):
        """设置日志记录器"""
        self.logger = logging.getLogger('PPTExtractor')
        self.logger.setLevel(logging.DEBUG)
        
        # 清除现有的处理器
        self.logger.handlers.clear()
        
        # 创建格式化器
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        
        # 控制台处理器
        if self.enable_console:
            console_handler = logging.StreamHandler(sys.stdout)
            console_handler.setLevel(logging.INFO)
            console_handler.setFormatter(formatter)
            self.logger.addHandler(console_handler)
        
        # 文件处理器
        if self.log_file:
            try:
                os.makedirs(os.path.dirname(self.log_file), exist_ok=True)
                file_handler = logging.FileHandler(self.log_file, encoding='utf-8')
                file_handler.setLevel(logging.DEBUG)
                file_handler.setFormatter(formatter)
                self.logger.addHandler(file_handler)
            except Exception as e:
                print(f"无法创建日志文件 {self.log_file}: {e}")
    
    def handle_error(self, error: Exception, context: Optional[Dict[str, Any]] = None) -> PPTExtractorError:
        """处理异常"""
        if isinstance(error, PPTExtractorError):
            ppt_error = error
        else:
            # 将标准异常转换为PPT提取器异常
            error_code = self._classify_error(error)
            ppt_error = PPTExtractorError(
                message=str(error),
                error_code=error_code,
                details={
                    'original_type': type(error).__name__,
                    'traceback': traceback.format_exc(),
                    'context': context or {}
                }
            )
        
        # 记录错误
        self.errors.append(ppt_error)
        
        # 记录日志
        self.logger.error(f"{ppt_error.error_code.value}: {ppt_error.message}")
        if ppt_error.details.get('traceback'):
            self.logger.debug(f"详细错误信息:\n{ppt_error.details['traceback']}")
        
        return ppt_error
    
    def handle_warning(self, message: str, context: Optional[Dict[str, Any]] = None):
        """处理警告"""
        warning = {
            'message': message,
            'timestamp': datetime.now(),
            'context': context or {}
        }
        
        self.warnings.append(warning)
        self.logger.warning(message)
    
    def _classify_error(self, error: Exception) -> ErrorCode:
        """分类错误类型"""
        error_type = type(error).__name__
        error_message = str(error).lower()
        
        # 文件相关错误
        if isinstance(error, FileNotFoundError):
            return ErrorCode.FILE_NOT_FOUND
        elif isinstance(error, PermissionError):
            if 'directory' in error_message or 'folder' in error_message:
                return ErrorCode.DIR_ACCESS_DENIED
            else:
                return ErrorCode.FILE_ACCESS_DENIED
        elif isinstance(error, IsADirectoryError):
            return ErrorCode.FILE_FORMAT_UNSUPPORTED
        
        # 内存相关错误
        elif isinstance(error, MemoryError):
            return ErrorCode.MEMORY_ERROR
        
        # 磁盘空间错误
        elif isinstance(error, OSError):
            if 'no space left' in error_message or 'disk full' in error_message:
                return ErrorCode.DISK_SPACE_ERROR
            elif 'permission denied' in error_message:
                return ErrorCode.PERMISSION_ERROR
            else:
                return ErrorCode.UNKNOWN_ERROR
        
        # ZIP相关错误（文件损坏）
        elif 'zipfile' in error_type.lower() or 'bad zip file' in error_message:
            return ErrorCode.FILE_CORRUPTED
        
        # 其他错误
        else:
            return ErrorCode.UNKNOWN_ERROR
    
    def get_error_summary(self) -> Dict[str, Any]:
        """获取错误摘要"""
        error_counts = {}
        for error in self.errors:
            code = error.error_code.value
            error_counts[code] = error_counts.get(code, 0) + 1
        
        return {
            'total_errors': len(self.errors),
            'total_warnings': len(self.warnings),
            'error_counts': error_counts,
            'latest_errors': [{
                'code': error.error_code.value,
                'message': error.message,
                'timestamp': error.timestamp.isoformat()
            } for error in self.errors[-5:]]  # 最近5个错误
        }
    
    def generate_error_report(self) -> str:
        """生成错误报告"""
        report = []
        report.append("PPT嵌入对象提取 - 错误报告")
        report.append("=" * 50)
        report.append(f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        report.append("")
        
        # 摘要信息
        summary = self.get_error_summary()
        report.append("摘要信息:")
        report.append(f"  总错误数: {summary['total_errors']}")
        report.append(f"  总警告数: {summary['total_warnings']}")
        report.append("")
        
        # 错误统计
        if summary['error_counts']:
            report.append("错误统计:")
            for code, count in summary['error_counts'].items():
                report.append(f"  {code}: {count} 次")
            report.append("")
        
        # 详细错误信息
        if self.errors:
            report.append("详细错误信息:")
            report.append("-" * 30)
            for i, error in enumerate(self.errors, 1):
                report.append(f"{i}. [{error.error_code.value}] {error.message}")
                report.append(f"   时间: {error.timestamp.strftime('%Y-%m-%d %H:%M:%S')}")
                if error.details.get('context'):
                    report.append(f"   上下文: {error.details['context']}")
                report.append("")
        
        # 警告信息
        if self.warnings:
            report.append("警告信息:")
            report.append("-" * 30)
            for i, warning in enumerate(self.warnings, 1):
                report.append(f"{i}. {warning['message']}")
                report.append(f"   时间: {warning['timestamp'].strftime('%Y-%m-%d %H:%M:%S')}")
                if warning.get('context'):
                    report.append(f"   上下文: {warning['context']}")
                report.append("")
        
        return "\n".join(report)
    
    def save_error_report(self, file_path: str):
        """保存错误报告到文件"""
        try:
            os.makedirs(os.path.dirname(file_path), exist_ok=True)
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(self.generate_error_report())
            self.logger.info(f"错误报告已保存到: {file_path}")
        except Exception as e:
            self.logger.error(f"保存错误报告失败: {e}")
    
    def clear_errors(self):
        """清除所有错误和警告"""
        self.errors.clear()
        self.warnings.clear()
        self.logger.info("错误和警告记录已清除")


def safe_execute(func, *args, error_handler: Optional[ErrorHandler] = None, **kwargs):
    """安全执行函数，自动处理异常"""
    try:
        return func(*args, **kwargs)
    except Exception as e:
        if error_handler:
            return error_handler.handle_error(e, {
                'function': func.__name__,
                'args': str(args)[:200],  # 限制长度
                'kwargs': str(kwargs)[:200]
            })
        else:
            raise


def validate_file_path(file_path: str, must_exist: bool = True) -> None:
    """验证文件路径"""
    if not file_path:
        raise PPTExtractorError("文件路径不能为空", ErrorCode.FILE_NOT_FOUND)
    
    if must_exist and not os.path.exists(file_path):
        raise PPTExtractorError(f"文件不存在: {file_path}", ErrorCode.FILE_NOT_FOUND)
    
    if must_exist and not os.path.isfile(file_path):
        raise PPTExtractorError(f"路径不是文件: {file_path}", ErrorCode.FILE_FORMAT_UNSUPPORTED)
    
    # 检查文件权限
    if must_exist:
        try:
            with open(file_path, 'rb') as f:
                f.read(1)  # 尝试读取一个字节
        except PermissionError:
            raise PPTExtractorError(f"没有读取文件的权限: {file_path}", ErrorCode.FILE_ACCESS_DENIED)
        except Exception as e:
            raise PPTExtractorError(f"文件访问错误: {e}", ErrorCode.FILE_CORRUPTED)


def validate_directory_path(dir_path: str, create_if_missing: bool = True) -> None:
    """验证目录路径"""
    if not dir_path:
        raise PPTExtractorError("目录路径不能为空", ErrorCode.DIR_NOT_FOUND)
    
    if os.path.exists(dir_path):
        if not os.path.isdir(dir_path):
            raise PPTExtractorError(f"路径不是目录: {dir_path}", ErrorCode.DIR_NOT_FOUND)
        
        # 检查写入权限
        if not os.access(dir_path, os.W_OK):
            raise PPTExtractorError(f"没有写入目录的权限: {dir_path}", ErrorCode.DIR_ACCESS_DENIED)
    
    elif create_if_missing:
        try:
            os.makedirs(dir_path, exist_ok=True)
        except PermissionError:
            raise PPTExtractorError(f"没有创建目录的权限: {dir_path}", ErrorCode.DIR_ACCESS_DENIED)
        except Exception as e:
            raise PPTExtractorError(f"创建目录失败: {e}", ErrorCode.DIR_CREATE_FAILED)
    else:
        raise PPTExtractorError(f"目录不存在: {dir_path}", ErrorCode.DIR_NOT_FOUND)


def check_disk_space(path: str, required_bytes: int) -> None:
    """检查磁盘空间"""
    try:
        import shutil
        free_bytes = shutil.disk_usage(path).free
        if free_bytes < required_bytes:
            raise PPTExtractorError(
                f"磁盘空间不足，需要 {required_bytes:,} 字节，可用 {free_bytes:,} 字节",
                ErrorCode.DISK_SPACE_ERROR
            )
    except Exception as e:
        if not isinstance(e, PPTExtractorError):
            raise PPTExtractorError(f"检查磁盘空间时出错: {e}", ErrorCode.UNKNOWN_ERROR)