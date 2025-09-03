#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
深度分析PPT文件中的原始文件名
"""

import zipfile
import xml.etree.ElementTree as ET
import os
import re
import struct

def analyze_ppt_for_original_names(ppt_path):
    """
    深度分析PPT文件，查找所有可能的原始文件名
    """
    print(f"正在分析PPT文件: {ppt_path}")
    print("="*60)
    
    with zipfile.ZipFile(ppt_path, 'r') as zip_file:
        # 1. 分析所有XML文件中的文本内容
        print("\n1. 分析XML文件中的文本内容:")
        print("-"*40)
        
        xml_files = [f for f in zip_file.namelist() if f.endswith('.xml')]
        for xml_file in xml_files:
            try:
                xml_data = zip_file.read(xml_file)
                root = ET.fromstring(xml_data)
                
                # 查找所有文本节点
                text_content = ET.tostring(root, encoding='unicode', method='text')
                
                # 查找可能的中文文件名（包含中文字符且有文件扩展名）
                chinese_pattern = r'[\u4e00-\u9fff]+.*?\.(xlsx?|docx?|pptx?|pdf|txt)'
                matches = re.findall(chinese_pattern, text_content, re.IGNORECASE)
                
                if matches:
                    print(f"  {xml_file}: 发现可能的中文文件名:")
                    for match in matches:
                        print(f"    - {match}")
                        
                # 查找所有name属性
                for elem in root.iter():
                    name_attr = elem.get('name')
                    if name_attr and ('.' in name_attr or re.search(r'[\u4e00-\u9fff]', name_attr)):
                        print(f"  {xml_file}: name属性: {name_attr}")
                        
            except Exception as e:
                print(f"  解析{xml_file}时出错: {e}")
        
        # 2. 分析关系文件
        print("\n2. 分析关系文件:")
        print("-"*40)
        
        rel_files = [f for f in zip_file.namelist() if f.endswith('.rels')]
        for rel_file in rel_files:
            try:
                rel_data = zip_file.read(rel_file)
                root = ET.fromstring(rel_data)
                
                print(f"  {rel_file}:")
                for relationship in root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                    rel_id = relationship.get('Id')
                    target = relationship.get('Target')
                    rel_type = relationship.get('Type')
                    
                    if target and ('embed' in target.lower() or 'object' in target.lower()):
                        print(f"    {rel_id}: {target} ({rel_type})")
                        
            except Exception as e:
                print(f"  解析{rel_file}时出错: {e}")
        
        # 3. 分析嵌入对象的二进制内容
        print("\n3. 分析嵌入对象的二进制内容:")
        print("-"*40)
        
        embed_files = [f for f in zip_file.namelist() if 'embed' in f.lower() or f.endswith('.bin')]
        for embed_file in embed_files[:5]:  # 只分析前5个
            try:
                embed_data = zip_file.read(embed_file)
                print(f"  {embed_file} ({len(embed_data)} bytes):")
                
                # 查找可能的文件名字符串
                # 尝试不同的编码
                for encoding in ['utf-16le', 'utf-8', 'gbk', 'utf-16be']:
                    try:
                        # 查找可能的文件名模式
                        text = embed_data.decode(encoding, errors='ignore')
                        
                        # 查找包含中文和文件扩展名的字符串
                        chinese_files = re.findall(r'[\u4e00-\u9fff][^\x00-\x1f]*?\.(xlsx?|docx?|pptx?|pdf|txt)', text, re.IGNORECASE)
                        if chinese_files:
                            print(f"    {encoding}编码发现: {chinese_files[:3]}")
                            
                        # 查找任何包含文件扩展名的字符串
                        file_patterns = re.findall(r'[^\x00-\x1f]{2,50}\.(xlsx?|docx?|pptx?|pdf|txt)', text, re.IGNORECASE)
                        if file_patterns:
                            print(f"    {encoding}编码文件名: {file_patterns[:3]}")
                            
                    except:
                        continue
                
                # 查找OLE复合文档头
                if embed_data.startswith(b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'):
                    print(f"    检测到OLE复合文档")
                    
                    # 尝试解析OLE目录结构
                    try:
                        # 简单的OLE解析
                        sector_size = 512
                        if len(embed_data) > 76:
                            # 读取目录流的位置
                            dir_first_sector = struct.unpack('<L', embed_data[48:52])[0]
                            if dir_first_sector < len(embed_data) // sector_size:
                                dir_offset = (dir_first_sector + 1) * sector_size
                                if dir_offset + 128 <= len(embed_data):
                                    # 读取目录项
                                    dir_entry = embed_data[dir_offset:dir_offset + 128]
                                    # 目录项名称（UTF-16LE编码）
                                    name_len = struct.unpack('<H', dir_entry[64:66])[0]
                                    if name_len > 0 and name_len <= 64:
                                        name_bytes = dir_entry[:name_len-2]  # 减去null终止符
                                        try:
                                            ole_name = name_bytes.decode('utf-16le')
                                            if ole_name and ole_name != '\x01Ole' and len(ole_name) > 1:
                                                print(f"    OLE根目录名: {ole_name}")
                                        except:
                                            pass
                    except Exception as e:
                        print(f"    OLE解析错误: {e}")
                        
            except Exception as e:
                print(f"  分析{embed_file}时出错: {e}")
        
        # 4. 查找docProps中的信息
        print("\n4. 分析文档属性:")
        print("-"*40)
        
        prop_files = [f for f in zip_file.namelist() if 'docProps' in f]
        for prop_file in prop_files:
            try:
                prop_data = zip_file.read(prop_file)
                root = ET.fromstring(prop_data)
                
                print(f"  {prop_file}:")
                # 查找所有文本内容
                for elem in root.iter():
                    if elem.text and (re.search(r'[\u4e00-\u9fff]', elem.text) or '.' in elem.text):
                        print(f"    {elem.tag}: {elem.text}")
                        
            except Exception as e:
                print(f"  解析{prop_file}时出错: {e}")

if __name__ == "__main__":
    ppt_file = "课程共建交付件清单和开发顺序0828 - 20250903145602.pptx"
    if os.path.exists(ppt_file):
        analyze_ppt_for_original_names(ppt_file)
    else:
        print(f"文件不存在: {ppt_file}")