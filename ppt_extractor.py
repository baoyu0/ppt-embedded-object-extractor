#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PPT嵌入对象提取工具
从PowerPoint文件中提取所有嵌入的Word、Excel和其他文件对象
"""

import os
import sys
import zipfile
import shutil
import argparse
from pathlib import Path
from typing import List, Dict, Tuple, Optional
import xml.etree.ElementTree as ET
from tqdm import tqdm
import chardet
from datetime import datetime

# 导入自定义模块
from file_type_detector import FileTypeDetector
from error_handler import ErrorHandler, PPTExtractorError, ErrorCode, validate_file_path, validate_directory_path, check_disk_space


class PPTExtractor:
    """PPT嵌入对象提取器"""
    
    def __init__(self, log_file: Optional[str] = None, enable_console_log: bool = True):
        self.extracted_files = []
        self.failed_files = []
        
        # 初始化组件
        self.file_detector = FileTypeDetector()
        self.error_handler = ErrorHandler(log_file, enable_console_log)
        
        # 统计信息
        self.stats = {
            'start_time': None,
            'end_time': None,
            'total_files_found': 0,
            'total_files_extracted': 0,
            'total_files_failed': 0,
            'total_size_extracted': 0
        }
    
    def extract_embedded_objects(self, ppt_path: str, output_dir: str) -> Dict[str, List[str]]:
        """
        从PPT文件中提取所有嵌入对象
        
        Args:
            ppt_path: PPT文件路径
            output_dir: 输出目录路径
            
        Returns:
            包含提取结果的字典
        """
        self.stats['start_time'] = datetime.now()
        
        try:
            # 验证输入参数
            validate_file_path(ppt_path, must_exist=True)
            validate_directory_path(output_dir, create_if_missing=True)
            
            # 检查文件格式
            if not ppt_path.lower().endswith(('.pptx', '.ppt')):
                raise PPTExtractorError(
                    "仅支持PPTX和PPT格式文件", 
                    ErrorCode.FILE_FORMAT_UNSUPPORTED
                )
            
            # 检查文件大小和磁盘空间
            file_size = os.path.getsize(ppt_path)
            check_disk_space(output_dir, file_size * 2)  # 预留2倍文件大小的空间
            
            self.error_handler.logger.info(f"开始分析PPT文件: {ppt_path}")
            self.error_handler.logger.info(f"文件大小: {self.file_detector.format_file_size(file_size)}")
            
            if ppt_path.lower().endswith('.pptx'):
                result = self._extract_from_pptx(ppt_path, output_dir)
            else:
                result = self._extract_from_ppt(ppt_path, output_dir)
            
            self.stats['end_time'] = datetime.now()
            return result
                
        except Exception as e:
            self.stats['end_time'] = datetime.now()
            error = self.error_handler.handle_error(e, {'ppt_path': ppt_path, 'output_dir': output_dir})
            return {
                'success': [],
                'failed': [{'file': ppt_path, 'error': str(error)}],
                'summary': {'total': 0, 'extracted': 0, 'failed': 1},
                'stats': self.stats
            }
    
    def _extract_from_pptx(self, pptx_path: str, output_dir: str) -> Dict[str, List[str]]:
        """
        从PPTX文件中提取嵌入对象
        """
        extracted_files = []
        failed_files = []
        
        try:
            with zipfile.ZipFile(pptx_path, 'r') as zip_file:
                # 获取所有文件列表
                file_list = zip_file.namelist()
                
                # 查找嵌入对象目录
                embedded_dirs = [
                    'ppt/embeddings/',
                    'ppt/media/',
                    'word/embeddings/',
                    'xl/embeddings/'
                ]
                
                embedded_files = []
                for file_path in file_list:
                    for embed_dir in embedded_dirs:
                        if file_path.startswith(embed_dir) and not file_path.endswith('/'):
                            embedded_files.append(file_path)
                
                # 分析关系文件以获取更多信息
                relationships = self._parse_relationships(zip_file)
                
                # 提取嵌入文件
                if embedded_files:
                    print(f"发现 {len(embedded_files)} 个嵌入对象")
                    
                    for file_path in tqdm(embedded_files, desc="提取嵌入对象"):
                        try:
                            # 提取文件
                            file_data = zip_file.read(file_path)
                            
                            # 确定文件类型和扩展名
                            file_ext, file_type, mime_type = self.file_detector.detect_file_type(file_data, file_path)
                            file_category = self.file_detector.get_file_category(file_ext[1:])  # 去掉点号
                            
                            # 生成输出文件名
                            base_name = os.path.basename(file_path)
                            if '.' not in base_name:
                                base_name += file_ext
                            
                            output_path = os.path.join(output_dir, base_name)
                            
                            # 处理重名文件
                            counter = 1
                            original_output_path = output_path
                            while os.path.exists(output_path):
                                name, ext = os.path.splitext(original_output_path)
                                output_path = f"{name}_{counter}{ext}"
                                counter += 1
                            
                            # 保存文件
                            try:
                                with open(output_path, 'wb') as output_file:
                                    output_file.write(file_data)
                                
                                # 验证文件是否正确保存
                                if not os.path.exists(output_path) or os.path.getsize(output_path) != len(file_data):
                                    raise PPTExtractorError("文件保存验证失败", ErrorCode.FILE_SAVE_FAILED)
                                
                                file_info = {
                                    'original_path': file_path,
                                    'output_path': output_path,
                                    'file_type': file_type,
                                    'file_category': file_category,
                                    'mime_type': mime_type,
                                    'size': len(file_data),
                                    'formatted_size': self.file_detector.format_file_size(len(file_data))
                                }
                                
                                extracted_files.append(file_info)
                                self.stats['total_files_extracted'] += 1
                                self.stats['total_size_extracted'] += len(file_data)
                                
                                self.error_handler.logger.info(
                                    f"提取成功: {os.path.basename(output_path)} ({file_type}, {file_info['formatted_size']})"
                                )
                                
                            except Exception as save_error:
                                raise PPTExtractorError(
                                    f"保存文件失败: {str(save_error)}", 
                                    ErrorCode.FILE_SAVE_FAILED
                                )
                            
                        except Exception as e:
                            self.stats['total_files_failed'] += 1
                            error_info = {
                                'file': file_path,
                                'error': str(e),
                                'error_type': type(e).__name__
                            }
                            failed_files.append(error_info)
                            self.error_handler.handle_error(e, {
                                'file_path': file_path,
                                'operation': 'extract_embedded_object'
                            })
                else:
                    print("未发现嵌入对象")
                
        except Exception as e:
            failed_files.append({
                'file': pptx_path,
                'error': f"无法打开PPTX文件: {str(e)}"
            })
        
        # 更新统计信息
        self.stats['total_files_found'] = len(extracted_files) + len(failed_files)
        
        return {
            'success': extracted_files,
            'failed': failed_files,
            'summary': {
                'total': len(extracted_files) + len(failed_files),
                'extracted': len(extracted_files),
                'failed': len(failed_files)
            },
            'stats': self.stats
        }
    
    def _extract_from_ppt(self, ppt_path: str, output_dir: str) -> Dict[str, List[str]]:
        """
        从PPT文件中提取嵌入对象（旧格式）
        注意：PPT格式的提取相对复杂，这里提供基础实现
        """
        print("警告：PPT格式文件的嵌入对象提取功能有限")
        print("建议将PPT文件转换为PPTX格式以获得更好的提取效果")
        
        return {
            'success': [],
            'failed': [{
                'file': ppt_path,
                'error': '暂不支持PPT格式文件的嵌入对象提取，请转换为PPTX格式'
            }],
            'summary': {'total': 1, 'extracted': 0, 'failed': 1}
        }
    
    def _parse_relationships(self, zip_file: zipfile.ZipFile) -> Dict[str, str]:
        """
        解析关系文件以获取嵌入对象信息
        """
        relationships = {}
        
        try:
            # 查找关系文件
            rel_files = [f for f in zip_file.namelist() if f.endswith('.rels')]
            
            for rel_file in rel_files:
                try:
                    rel_data = zip_file.read(rel_file)
                    root = ET.fromstring(rel_data)
                    
                    # 解析关系
                    for relationship in root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                        rel_id = relationship.get('Id')
                        target = relationship.get('Target')
                        rel_type = relationship.get('Type')
                        
                        if rel_id and target:
                            relationships[rel_id] = {
                                'target': target,
                                'type': rel_type
                            }
                            
                except Exception as e:
                    print(f"解析关系文件 {rel_file} 时出错: {str(e)}")
                    
        except Exception as e:
            print(f"解析关系文件时出错: {str(e)}")
        
        return relationships
    

    

    
    def print_extraction_report(self, result: Dict[str, List[str]]):
        """
        打印提取报告
        """
        print("\n" + "="*60)
        print("PPT嵌入对象提取报告")
        print("="*60)
        
        # 基本统计信息
        summary = result.get('summary', {})
        stats = result.get('stats', {})
        
        print(f"总计发现文件: {summary.get('total', 0)}")
        print(f"成功提取: {summary.get('extracted', 0)}")
        print(f"提取失败: {summary.get('failed', 0)}")
        
        if stats:
            if stats.get('start_time') and stats.get('end_time'):
                duration = stats['end_time'] - stats['start_time']
                print(f"处理时间: {duration.total_seconds():.2f} 秒")
            
            if stats.get('total_size_extracted', 0) > 0:
                total_size = self.file_detector.format_file_size(stats['total_size_extracted'])
                print(f"总提取大小: {total_size}")
        
        # 成功提取的文件
        success_files = result.get('success', [])
        if success_files:
            print("\n✅ 成功提取的文件:")
            print("-" * 40)
            
            # 按文件类别分组
            by_category = {}
            for file_info in success_files:
                category = file_info.get('file_category', '其他')
                if category not in by_category:
                    by_category[category] = []
                by_category[category].append(file_info)
            
            for category, files in by_category.items():
                print(f"\n📁 {category} ({len(files)} 个文件):")
                for file_info in files:
                    print(f"  ✓ {os.path.basename(file_info['output_path'])}")
                    print(f"    类型: {file_info['file_type']}")
                    print(f"    大小: {file_info['formatted_size']}")
                    print(f"    保存路径: {file_info['output_path']}")
                    print()
        
        # 失败的文件
        failed_files = result.get('failed', [])
        if failed_files:
            print("\n❌ 提取失败的文件:")
            print("-" * 40)
            for file_info in failed_files:
                print(f"  ✗ {file_info.get('file', 'Unknown')}")
                print(f"    错误类型: {file_info.get('error_type', 'Unknown')}")
                print(f"    错误信息: {file_info.get('error', 'Unknown error')}")
                print()
        
        # 错误摘要
        if hasattr(self, 'error_handler') and hasattr(self.error_handler, 'get_error_summary'):
            try:
                error_summary = self.error_handler.get_error_summary()
                if error_summary:
                    print("\n📊 错误摘要:")
                    print("-" * 40)
                    for error_type, count in error_summary.items():
                        print(f"  {error_type}: {count} 次")
            except:
                pass
        
        print("\n" + "="*60)


def main():
    """
    主程序入口
    """
    print("PPT嵌入对象提取工具")
    print("=" * 30)
    
    # 获取用户输入
    if len(sys.argv) > 1:
        ppt_path = sys.argv[1]
    else:
        ppt_path = input("请输入PPT文件路径: ").strip('"')
    
    if len(sys.argv) > 2:
        output_dir = sys.argv[2]
    else:
        output_dir = input("请输入输出目录路径 (回车使用默认目录 'extracted_files'): ").strip('"')
        if not output_dir:
            output_dir = "extracted_files"
    
    # 创建提取器实例
    extractor = PPTExtractor()
    
    # 执行提取
    result = extractor.extract_embedded_objects(ppt_path, output_dir)
    
    # 显示报告
    extractor.print_extraction_report(result)
    
    print("\n提取完成！")


if __name__ == "__main__":
    main()