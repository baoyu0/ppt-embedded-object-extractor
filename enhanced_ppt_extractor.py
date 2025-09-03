#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
增强版PPT嵌入对象提取器
支持用户自定义文件名映射，解决原始中文文件名丢失问题
"""

import os
import zipfile
import xml.etree.ElementTree as ET
import json
from pathlib import Path
from file_type_detector import FileTypeDetector
from error_handler import ErrorHandler

class EnhancedPPTExtractor:
    def __init__(self, output_dir="extracted_objects_enhanced"):
        self.output_dir = output_dir
        self.file_detector = FileTypeDetector()
        self.error_handler = ErrorHandler()
        self.extracted_files = []
        self.failed_files = []
        
        # 创建输出目录
        os.makedirs(self.output_dir, exist_ok=True)
    
    def create_filename_mapping_template(self, ppt_file):
        """
        创建文件名映射模板，用户可以手动填写原始文件名
        """
        print(f"\n=== 创建文件名映射模板 ===")
        
        mapping_file = f"{Path(ppt_file).stem}_filename_mapping.json"
        
        try:
            with zipfile.ZipFile(ppt_file, 'r') as zip_ref:
                file_list = zip_ref.namelist()
                
                # 查找所有嵌入对象
                embedding_files = [f for f in file_list if 'embeddings' in f.lower()]
                
                # 创建映射模板
                mapping_template = {
                    "说明": "请在'original_name'字段中填写每个文件的原始中文名称",
                    "ppt_file": ppt_file,
                    "mappings": []
                }
                
                for i, emb_file in enumerate(embedding_files):
                    # 检测文件类型
                    try:
                        file_content = zip_ref.read(emb_file)
                        file_type_info = self.file_detector.detect_file_type(file_content)
                        
                        mapping_entry = {
                            "index": i + 1,
                            "embedded_path": emb_file,
                            "detected_type": file_type_info['type'],
                            "detected_extension": file_type_info['extension'],
                            "current_name": os.path.basename(emb_file),
                            "original_name": "",  # 用户需要填写
                            "description": ""  # 可选描述
                        }
                        
                        mapping_template["mappings"].append(mapping_entry)
                        
                    except Exception as e:
                        print(f"处理文件 {emb_file} 时出错: {e}")
                
                # 保存映射模板
                with open(mapping_file, 'w', encoding='utf-8') as f:
                    json.dump(mapping_template, f, ensure_ascii=False, indent=2)
                
                print(f"\n文件名映射模板已创建: {mapping_file}")
                print(f"找到 {len(mapping_template['mappings'])} 个嵌入对象")
                print("\n请编辑此文件，在'original_name'字段中填写原始文件名，然后使用extract_with_mapping()方法提取")
                
                return mapping_file
                
        except Exception as e:
            print(f"创建映射模板失败: {e}")
            return None
    
    def extract_with_mapping(self, ppt_file, mapping_file):
        """
        使用文件名映射提取嵌入对象
        """
        print(f"\n=== 使用映射文件提取嵌入对象 ===")
        
        if not os.path.exists(mapping_file):
            print(f"映射文件不存在: {mapping_file}")
            return False
        
        try:
            # 读取映射文件
            with open(mapping_file, 'r', encoding='utf-8') as f:
                mapping_data = json.load(f)
            
            print(f"读取映射文件: {mapping_file}")
            print(f"映射条目数: {len(mapping_data['mappings'])}")
            
            with zipfile.ZipFile(ppt_file, 'r') as zip_ref:
                for mapping in mapping_data['mappings']:
                    embedded_path = mapping['embedded_path']
                    original_name = mapping.get('original_name', '').strip()
                    detected_extension = mapping['detected_extension']
                    current_name = mapping['current_name']
                    
                    try:
                        # 读取文件内容
                        file_content = zip_ref.read(embedded_path)
                        
                        # 确定输出文件名
                        if original_name:
                            # 使用用户提供的原始名称
                            if not original_name.endswith(detected_extension):
                                output_filename = f"{original_name}{detected_extension}"
                            else:
                                output_filename = original_name
                        else:
                            # 使用当前名称或生成描述性名称
                            if current_name.endswith('.bin'):
                                base_name = Path(current_name).stem
                                output_filename = f"{base_name}{detected_extension}"
                            else:
                                output_filename = current_name
                        
                        # 处理重名文件
                        output_path = os.path.join(self.output_dir, output_filename)
                        counter = 1
                        while os.path.exists(output_path):
                            name_part = Path(output_filename).stem
                            ext_part = Path(output_filename).suffix
                            output_filename = f"{name_part}_{counter}{ext_part}"
                            output_path = os.path.join(self.output_dir, output_filename)
                            counter += 1
                        
                        # 保存文件
                        with open(output_path, 'wb') as output_file:
                            output_file.write(file_content)
                        
                        # 记录成功提取的文件
                        file_info = {
                            'original_path': embedded_path,
                            'output_path': output_path,
                            'output_filename': output_filename,
                            'original_name': original_name if original_name else "未指定",
                            'file_type': mapping['detected_type'],
                            'file_size': len(file_content)
                        }
                        
                        self.extracted_files.append(file_info)
                        print(f"✓ 提取成功: {output_filename} (原始名称: {file_info['original_name']})")
                        
                    except Exception as e:
                        error_info = {
                            'file': embedded_path,
                            'error': str(e)
                        }
                        self.failed_files.append(error_info)
                        print(f"✗ 提取失败: {embedded_path} - {e}")
            
            self.print_extraction_report()
            return True
            
        except Exception as e:
            print(f"使用映射文件提取失败: {e}")
            return False
    
    def extract_with_smart_naming(self, ppt_file):
        """
        使用智能命名策略提取嵌入对象
        """
        print(f"\n=== 使用智能命名策略提取嵌入对象 ===")
        
        try:
            with zipfile.ZipFile(ppt_file, 'r') as zip_ref:
                file_list = zip_ref.namelist()
                embedding_files = [f for f in file_list if 'embeddings' in f.lower()]
                
                print(f"找到 {len(embedding_files)} 个嵌入对象")
                
                # 按文件类型分组计数
                type_counters = {}
                
                for emb_file in embedding_files:
                    try:
                        file_content = zip_ref.read(emb_file)
                        file_type_info = self.file_detector.detect_file_type(file_content)
                        file_type = file_type_info['type']
                        extension = file_type_info['extension']
                        
                        # 智能命名策略
                        if file_type not in type_counters:
                            type_counters[file_type] = 1
                        else:
                            type_counters[file_type] += 1
                        
                        # 生成智能文件名
                        if file_type == 'Microsoft Word Document':
                            base_name = f"Word文档_{type_counters[file_type]:02d}"
                        elif file_type == 'Microsoft Excel Worksheet':
                            base_name = f"Excel表格_{type_counters[file_type]:02d}"
                        elif file_type == 'Microsoft PowerPoint Presentation':
                            base_name = f"PPT演示文稿_{type_counters[file_type]:02d}"
                        elif 'image' in file_type.lower():
                            base_name = f"图片_{type_counters[file_type]:02d}"
                        else:
                            base_name = f"{file_type}_{type_counters[file_type]:02d}"
                        
                        output_filename = f"{base_name}{extension}"
                        
                        # 处理重名文件
                        output_path = os.path.join(self.output_dir, output_filename)
                        counter = 1
                        while os.path.exists(output_path):
                            output_filename = f"{base_name}_副本{counter}{extension}"
                            output_path = os.path.join(self.output_dir, output_filename)
                            counter += 1
                        
                        # 保存文件
                        with open(output_path, 'wb') as output_file:
                            output_file.write(file_content)
                        
                        # 记录成功提取的文件
                        file_info = {
                            'original_path': emb_file,
                            'output_path': output_path,
                            'output_filename': output_filename,
                            'original_name': "智能命名",
                            'file_type': file_type,
                            'file_size': len(file_content)
                        }
                        
                        self.extracted_files.append(file_info)
                        print(f"✓ 提取成功: {output_filename}")
                        
                    except Exception as e:
                        error_info = {
                            'file': emb_file,
                            'error': str(e)
                        }
                        self.failed_files.append(error_info)
                        print(f"✗ 提取失败: {emb_file} - {e}")
                
                self.print_extraction_report()
                return True
                
        except Exception as e:
            print(f"智能命名提取失败: {e}")
            return False
    
    def print_extraction_report(self):
        """
        打印提取报告
        """
        print(f"\n=== 提取报告 ===")
        print(f"成功提取: {len(self.extracted_files)} 个文件")
        print(f"提取失败: {len(self.failed_files)} 个文件")
        print(f"输出目录: {self.output_dir}")
        
        if self.extracted_files:
            print("\n成功提取的文件:")
            for file_info in self.extracted_files:
                print(f"  {file_info['output_filename']} ({file_info['file_type']}, {file_info['file_size']} bytes)")
                if file_info['original_name'] != "智能命名" and file_info['original_name'] != "未指定":
                    print(f"    原始名称: {file_info['original_name']}")
        
        if self.failed_files:
            print("\n提取失败的文件:")
            for error_info in self.failed_files:
                print(f"  {error_info['file']}: {error_info['error']}")

def main():
    """
    主函数 - 提供多种提取选项
    """
    ppt_file = "课程共建交付件清单和开发顺序0828 - 20250903145602.pptx"
    
    if not os.path.exists(ppt_file):
        print(f"PPT文件不存在: {ppt_file}")
        return
    
    extractor = EnhancedPPTExtractor()
    
    print("=== 增强版PPT嵌入对象提取器 ===")
    print("\n选择提取方式:")
    print("1. 创建文件名映射模板（推荐）")
    print("2. 使用现有映射文件提取")
    print("3. 智能命名提取")
    
    choice = input("\n请选择 (1-3): ").strip()
    
    if choice == '1':
        # 创建映射模板
        mapping_file = extractor.create_filename_mapping_template(ppt_file)
        if mapping_file:
            print(f"\n请编辑 {mapping_file} 文件，填写原始文件名后，")
            print("再次运行此程序并选择选项2进行提取。")
    
    elif choice == '2':
        # 使用映射文件提取
        mapping_file = f"{Path(ppt_file).stem}_filename_mapping.json"
        if os.path.exists(mapping_file):
            extractor.extract_with_mapping(ppt_file, mapping_file)
        else:
            print(f"映射文件不存在: {mapping_file}")
            print("请先选择选项1创建映射模板")
    
    elif choice == '3':
        # 智能命名提取
        extractor.extract_with_smart_naming(ppt_file)
    
    else:
        print("无效选择")

if __name__ == "__main__":
    main()