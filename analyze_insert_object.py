#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
专门分析通过"插入对象"功能插入的文件名提取脚本
针对用户确认的插入方式：插入 → 对象 → 由文件创建
"""

import zipfile
import xml.etree.ElementTree as ET
import os
import re
from pathlib import Path

def analyze_insert_object_names(ppt_file):
    """
    专门分析通过插入对象功能插入的文件的原始名称
    """
    print(f"\n=== 分析插入对象文件名：{ppt_file} ===")
    
    if not os.path.exists(ppt_file):
        print(f"错误：文件不存在 {ppt_file}")
        return
    
    try:
        with zipfile.ZipFile(ppt_file, 'r') as zip_ref:
            # 获取所有文件列表
            file_list = zip_ref.namelist()
            print(f"\n文件总数：{len(file_list)}")
            
            # 1. 分析关系文件中的目标信息
            print("\n=== 1. 分析关系文件 ===")
            rel_files = [f for f in file_list if f.endswith('.rels')]
            object_relations = {}
            
            for rel_file in rel_files:
                print(f"\n检查关系文件：{rel_file}")
                try:
                    content = zip_ref.read(rel_file).decode('utf-8')
                    root = ET.fromstring(content)
                    
                    # 查找所有关系
                    for relationship in root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                        rel_id = relationship.get('Id', '')
                        target = relationship.get('Target', '')
                        rel_type = relationship.get('Type', '')
                        
                        # 重点关注嵌入对象关系
                        if 'oleObject' in target.lower() or 'embeddings' in target.lower():
                            object_relations[rel_id] = {
                                'target': target,
                                'type': rel_type,
                                'rel_file': rel_file
                            }
                            print(f"  找到对象关系 - ID: {rel_id}, Target: {target}, Type: {rel_type}")
                            
                except Exception as e:
                    print(f"  解析关系文件失败：{e}")
            
            # 2. 分析幻灯片XML中的对象信息
            print("\n=== 2. 分析幻灯片XML中的对象信息 ===")
            slide_files = [f for f in file_list if f.startswith('ppt/slides/slide') and f.endswith('.xml')]
            
            for slide_file in slide_files:
                print(f"\n检查幻灯片：{slide_file}")
                try:
                    content = zip_ref.read(slide_file).decode('utf-8')
                    root = ET.fromstring(content)
                    
                    # 定义命名空间
                    namespaces = {
                        'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
                        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
                        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
                    }
                    
                    # 查找所有OLE对象
                    ole_objects = root.findall('.//p:oleObj', namespaces)
                    print(f"  找到 {len(ole_objects)} 个OLE对象")
                    
                    for i, ole_obj in enumerate(ole_objects):
                        print(f"\n  --- OLE对象 {i+1} ---")
                        
                        # 获取关系ID
                        rel_id = ole_obj.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id', '')
                        print(f"    关系ID: {rel_id}")
                        
                        # 获取所有属性
                        for attr_name, attr_value in ole_obj.attrib.items():
                            print(f"    属性 {attr_name}: {attr_value}")
                        
                        # 查找相关的关系信息
                        if rel_id in object_relations:
                            rel_info = object_relations[rel_id]
                            print(f"    关联文件: {rel_info['target']}")
                            print(f"    关系类型: {rel_info['type']}")
                    
                    # 3. 深度搜索所有可能包含原始文件名的元素
                    print(f"\n  --- 深度搜索文件名信息 ---")
                    
                    # 搜索所有cNvPr元素（通常包含名称信息）
                    cnv_pr_elements = root.findall('.//a:cNvPr', namespaces)
                    print(f"  找到 {len(cnv_pr_elements)} 个cNvPr元素")
                    
                    for j, cnv_pr in enumerate(cnv_pr_elements):
                        name = cnv_pr.get('name', '')
                        descr = cnv_pr.get('descr', '')
                        title = cnv_pr.get('title', '')
                        
                        if name or descr or title:
                            print(f"    cNvPr {j+1}:")
                            if name:
                                print(f"      名称: {name}")
                            if descr:
                                print(f"      描述: {descr}")
                            if title:
                                print(f"      标题: {title}")
                    
                    # 搜索所有包含中文或文件扩展名的属性
                    all_text = ET.tostring(root, encoding='unicode')
                    
                    # 查找可能的文件名模式
                    filename_patterns = [
                        r'name="([^"]*[\u4e00-\u9fff][^"]*?)"',  # name属性中的中文
                        r'title="([^"]*[\u4e00-\u9fff][^"]*?)"',  # title属性中的中文
                        r'descr="([^"]*[\u4e00-\u9fff][^"]*?)"',  # descr属性中的中文
                        r'([\u4e00-\u9fff][^<>"]*?\.(docx?|xlsx?|pptx?|pdf|txt))',  # 中文文件名
                        r'"([^"]*[\u4e00-\u9fff][^"]*\.(docx?|xlsx?|pptx?|pdf|txt))"',  # 引号中的中文文件名
                    ]
                    
                    found_names = set()
                    for pattern in filename_patterns:
                        matches = re.findall(pattern, all_text, re.IGNORECASE)
                        if matches:
                            print(f"\n    模式匹配结果:")
                            for match in matches:
                                if isinstance(match, tuple):
                                    match = match[0]  # 取第一个捕获组
                                if match and match not in found_names:
                                    found_names.add(match)
                                    print(f"      *** 可能的原始文件名: {match} ***")
                    
                except Exception as e:
                    print(f"  解析幻灯片失败：{e}")
            
            # 4. 检查嵌入对象目录
            print("\n=== 3. 检查嵌入对象目录 ===")
            embedding_files = [f for f in file_list if 'embeddings' in f.lower()]
            if embedding_files:
                print(f"找到 {len(embedding_files)} 个嵌入文件：")
                for emb_file in embedding_files:
                    print(f"  {emb_file}")
                    
                    # 尝试从嵌入文件本身提取原始名称
                    if not emb_file.endswith('.bin'):
                        try:
                            # 对于非.bin文件，检查其内部是否有原始文件名信息
                            file_content = zip_ref.read(emb_file)
                            
                            # 尝试解析为Office文档
                            if emb_file.endswith(('.docx', '.xlsx', '.pptx')):
                                print(f"    尝试解析Office文档: {emb_file}")
                                # 这里可以进一步解析嵌入的Office文档
                                
                        except Exception as e:
                            print(f"    解析嵌入文件失败: {e}")
            else:
                print("未找到embeddings目录")
            
            # 5. 检查所有XML文件中的文本内容
            print("\n=== 4. 全局搜索中文文件名 ===")
            xml_files = [f for f in file_list if f.endswith('.xml')]
            
            all_found_names = set()
            for xml_file in xml_files:
                try:
                    content = zip_ref.read(xml_file).decode('utf-8')
                    
                    # 搜索所有可能的中文文件名
                    chinese_filename_patterns = [
                        r'([\u4e00-\u9fff][^<>"\s]*\.(docx?|xlsx?|pptx?|pdf|txt))',
                        r'"([^"]*[\u4e00-\u9fff][^"]*\.(docx?|xlsx?|pptx?|pdf|txt))"',
                        r'>([^<]*[\u4e00-\u9fff][^<]*\.(docx?|xlsx?|pptx?|pdf|txt))<',
                    ]
                    
                    for pattern in chinese_filename_patterns:
                        matches = re.findall(pattern, content, re.IGNORECASE)
                        for match in matches:
                            if isinstance(match, tuple):
                                match = match[0]
                            if match and len(match) > 3:  # 过滤太短的匹配
                                all_found_names.add(match)
                                
                except Exception as e:
                    continue
            
            if all_found_names:
                print("\n*** 全局搜索发现的可能文件名: ***")
                for name in sorted(all_found_names):
                    print(f"  {name}")
            else:
                print("\n未在XML文件中发现中文文件名")
            
            # 6. 检查docProps中的应用程序特定信息
            print("\n=== 5. 检查docProps中的详细信息 ===")
            doc_props_files = [f for f in file_list if f.startswith('docProps/') and f.endswith('.xml')]
            
            for prop_file in doc_props_files:
                print(f"\n检查属性文件：{prop_file}")
                try:
                    content = zip_ref.read(prop_file).decode('utf-8')
                    
                    # 查找所有包含中文的内容
                    chinese_matches = re.findall(r'[\u4e00-\u9fff]+[^<>]*', content)
                    if chinese_matches:
                        print(f"  找到中文内容：")
                        for match in chinese_matches[:10]:
                            print(f"    {match.strip()}")
                    
                    # 查找可能的文件名引用
                    filename_refs = re.findall(r'[^<>]*\.(docx?|xlsx?|pptx?|pdf|txt)[^<>]*', content, re.IGNORECASE)
                    if filename_refs:
                        print(f"  找到文件名引用：")
                        for ref in filename_refs:
                            print(f"    {ref.strip()}")
                            
                except Exception as e:
                    print(f"  解析属性文件失败：{e}")
            
            print("\n=== 分析完成 ===")
            
    except Exception as e:
        print(f"分析失败：{e}")

if __name__ == "__main__":
    # 测试文件
    test_file = "课程共建交付件清单和开发顺序0828 - 20250903145602.pptx"
    
    if os.path.exists(test_file):
        analyze_insert_object_names(test_file)
    else:
        print(f"测试文件不存在：{test_file}")
        print("请将PPT文件放在脚本同目录下，或修改文件路径")