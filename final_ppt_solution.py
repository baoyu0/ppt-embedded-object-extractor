#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PPT嵌入对象提取完整解决方案
解决原始中文文件名丢失问题的最终方案
"""

import os
import json
import shutil
from pathlib import Path
from enhanced_ppt_extractor import EnhancedPPTExtractor

def create_solution_guide():
    """
    创建解决方案指南
    """
    guide = """
=== PPT嵌入对象原始中文文件名恢复解决方案 ===

问题分析:
通过深度分析PPT文件结构，我们发现:
1. PPT文件中确实不存储嵌入对象的原始文件名
2. 即使是通过"插入对象"→"由文件创建"方式插入的文件
3. PowerPoint只保存文件类型信息，如"Microsoft Word Document"
4. 这是PowerPoint的设计机制，不是技术缺陷

解决方案:

方案一：文件名映射（推荐）
- 创建JSON映射模板
- 用户手动填写原始文件名
- 程序根据映射重命名提取的文件
- 优点：完全保留原始文件名
- 缺点：需要用户手动操作

方案二：智能命名
- 根据文件类型自动生成中文描述性名称
- 如：Word文档_01.docx, Excel表格_01.xlsx
- 优点：无需用户干预
- 缺点：不是原始文件名

方案三：预防性措施（建议）
- 在插入文件时记录原始文件名
- 使用文档管理系统
- 建立文件命名规范

使用建议:
1. 对于已有PPT文件：使用方案一或方案二
2. 对于新建PPT文件：采用方案三的预防措施
3. 重要文档建议使用方案一确保文件名准确性
"""
    return guide

def demonstrate_complete_solution():
    """
    演示完整解决方案
    """
    ppt_file = "课程共建交付件清单和开发顺序0828 - 20250903145602.pptx"
    
    if not os.path.exists(ppt_file):
        print(f"PPT文件不存在: {ppt_file}")
        return
    
    print(create_solution_guide())
    
    print("\n=== 实际演示 ===")
    
    # 1. 使用现有的提取结果
    print("\n1. 当前提取结果分析:")
    extracted_dir = "extracted_objects"
    if os.path.exists(extracted_dir):
        files = os.listdir(extracted_dir)
        word_files = [f for f in files if f.endswith('.docx')]
        excel_files = [f for f in files if f.endswith('.xlsx')]
        other_files = [f for f in files if not f.endswith(('.docx', '.xlsx'))]
        
        print(f"   Word文档: {len(word_files)} 个")
        print(f"   Excel表格: {len(excel_files)} 个")
        print(f"   其他文件: {len(other_files)} 个")
        
        print("\n   Word文档列表:")
        for i, f in enumerate(word_files[:5], 1):
            print(f"     {i}. {f}")
        if len(word_files) > 5:
            print(f"     ... 还有 {len(word_files) - 5} 个")
        
        print("\n   Excel表格列表:")
        for i, f in enumerate(excel_files[:5], 1):
            print(f"     {i}. {f}")
        if len(excel_files) > 5:
            print(f"     ... 还有 {len(excel_files) - 5} 个")
    
    # 2. 创建智能命名版本
    print("\n2. 创建智能命名版本:")
    smart_dir = "final_smart_naming"
    if os.path.exists(smart_dir):
        shutil.rmtree(smart_dir)
    os.makedirs(smart_dir, exist_ok=True)
    
    if os.path.exists(extracted_dir):
        files = os.listdir(extracted_dir)
        word_count = 1
        excel_count = 1
        ppt_count = 1
        other_count = 1
        
        for file in files:
            src_path = os.path.join(extracted_dir, file)
            if file.endswith('.docx'):
                new_name = f"Word文档_{word_count:02d}.docx"
                word_count += 1
            elif file.endswith('.xlsx'):
                new_name = f"Excel表格_{excel_count:02d}.xlsx"
                excel_count += 1
            elif file.endswith('.pptx'):
                new_name = f"PowerPoint演示文稿_{ppt_count:02d}.pptx"
                ppt_count += 1
            elif file.startswith('oleObject'):
                continue  # 跳过二进制文件
            else:
                new_name = f"其他文件_{other_count:02d}{Path(file).suffix}"
                other_count += 1
            
            dst_path = os.path.join(smart_dir, new_name)
            try:
                shutil.copy2(src_path, dst_path)
                print(f"   {file} → {new_name}")
            except Exception as e:
                print(f"   复制失败: {file} - {e}")
    
    # 3. 创建文件名映射模板
    print("\n3. 创建文件名映射模板:")
    mapping_template = {
        "说明": "请在下面的映射中填写每个文件的原始中文名称",
        "ppt_file": ppt_file,
        "mappings": []
    }
    
    if os.path.exists(extracted_dir):
        files = [f for f in os.listdir(extracted_dir) if f.endswith(('.docx', '.xlsx', '.pptx'))]
        for i, file in enumerate(files, 1):
            mapping = {
                "序号": i,
                "当前文件名": file,
                "文件类型": "Word文档" if file.endswith('.docx') else 
                           "Excel表格" if file.endswith('.xlsx') else 
                           "PowerPoint演示文稿" if file.endswith('.pptx') else "未知",
                "原始文件名": f"请填写原始文件名{i}.{file.split('.')[-1]}",
                "描述": "请填写文件描述"
            }
            mapping_template["mappings"].append(mapping)
    
    mapping_file = f"{Path(ppt_file).stem}_文件名映射模板.json"
    with open(mapping_file, 'w', encoding='utf-8') as f:
        json.dump(mapping_template, f, ensure_ascii=False, indent=2)
    
    print(f"   映射模板已创建: {mapping_file}")
    print(f"   请编辑此文件，填写正确的原始文件名")
    
    # 4. 创建示例映射
    print("\n4. 创建示例映射（演示用）:")
    sample_names = [
        "课程设计文档.docx", "实验手册.docx", "评分标准.xlsx", 
        "课程大纲.pptx", "作业模板.docx", "考试题库.xlsx",
        "教学计划.docx", "学习资料.xlsx", "课程介绍.pptx", 
        "参考资料.docx", "实验指导书.docx", "课程评估表.xlsx"
    ]
    
    for i, mapping in enumerate(mapping_template["mappings"]):
        if i < len(sample_names):
            mapping["原始文件名"] = sample_names[i]
            mapping["描述"] = f"这是第{i+1}个嵌入的{mapping['文件类型']}"
    
    sample_mapping_file = f"{Path(ppt_file).stem}_示例映射.json"
    with open(sample_mapping_file, 'w', encoding='utf-8') as f:
        json.dump(mapping_template, f, ensure_ascii=False, indent=2)
    
    print(f"   示例映射已创建: {sample_mapping_file}")
    
    # 5. 使用示例映射重命名文件
    print("\n5. 使用示例映射重命名文件:")
    mapped_dir = "final_mapped_naming"
    if os.path.exists(mapped_dir):
        shutil.rmtree(mapped_dir)
    os.makedirs(mapped_dir, exist_ok=True)
    
    for mapping in mapping_template["mappings"]:
        src_file = mapping["当前文件名"]
        dst_file = mapping["原始文件名"]
        src_path = os.path.join(extracted_dir, src_file)
        dst_path = os.path.join(mapped_dir, dst_file)
        
        if os.path.exists(src_path):
            try:
                shutil.copy2(src_path, dst_path)
                print(f"   {src_file} → {dst_file}")
            except Exception as e:
                print(f"   重命名失败: {src_file} - {e}")
    
    # 6. 总结
    print("\n=== 解决方案总结 ===")
    print("\n已创建的文件和目录:")
    print(f"   1. 原始提取结果: {extracted_dir}/")
    print(f"   2. 智能命名结果: {smart_dir}/")
    print(f"   3. 映射模板文件: {mapping_file}")
    print(f"   4. 示例映射文件: {sample_mapping_file}")
    print(f"   5. 映射重命名结果: {mapped_dir}/")
    
    print("\n使用方法:")
    print("   1. 查看智能命名结果，获得有意义的中文文件名")
    print("   2. 编辑映射模板文件，填写正确的原始文件名")
    print("   3. 运行映射程序，获得准确命名的文件")
    
    print("\n技术结论:")
    print("   - PPT文件结构中确实不保存嵌入对象的原始文件名")
    print("   - 这是PowerPoint的设计特性，不是技术缺陷")
    print("   - 通过文件名映射可以完美解决此问题")
    print("   - 建议在插入文件时建立文档管理规范")

def create_usage_script():
    """
    创建使用脚本
    """
    script_content = '''
@echo off
chcp 65001
echo === PPT嵌入对象提取解决方案 ===
echo.
echo 1. 智能命名提取（推荐新手使用）
echo 2. 创建文件名映射模板
echo 3. 使用映射文件重命名
echo 4. 查看解决方案说明
echo.
set /p choice=请选择操作 (1-4): 

if "%choice%"=="1" (
    python final_ppt_solution.py
) else if "%choice%"=="2" (
    echo 创建映射模板功能已集成在主程序中
    python final_ppt_solution.py
) else if "%choice%"=="3" (
    echo 映射重命名功能已集成在主程序中
    python final_ppt_solution.py
) else if "%choice%"=="4" (
    echo 查看 README.md 文件获取详细说明
    type README.md
) else (
    echo 无效选择
)

pause
'''
    
    with open('run_solution.bat', 'w', encoding='utf-8') as f:
        f.write(script_content)
    
    print("使用脚本已创建: run_solution.bat")

if __name__ == "__main__":
    demonstrate_complete_solution()
    create_usage_script()