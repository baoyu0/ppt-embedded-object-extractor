#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
增强版PPT提取器演示脚本
自动演示文件名映射模板创建和智能命名功能
"""

import os
import json
from enhanced_ppt_extractor import EnhancedPPTExtractor
from pathlib import Path

def demo_enhanced_extractor():
    """
    演示增强版提取器的功能
    """
    ppt_file = "课程共建交付件清单和开发顺序0828 - 20250903145602.pptx"
    
    if not os.path.exists(ppt_file):
        print(f"PPT文件不存在: {ppt_file}")
        return
    
    print("=== 增强版PPT嵌入对象提取器演示 ===")
    
    # 1. 创建文件名映射模板
    print("\n步骤1: 创建文件名映射模板")
    extractor = EnhancedPPTExtractor("extracted_objects_enhanced")
    mapping_file = extractor.create_filename_mapping_template(ppt_file)
    
    if mapping_file and os.path.exists(mapping_file):
        print(f"\n映射模板已创建: {mapping_file}")
        
        # 读取并显示模板内容
        with open(mapping_file, 'r', encoding='utf-8') as f:
            mapping_data = json.load(f)
        
        print("\n模板内容预览:")
        for i, mapping in enumerate(mapping_data['mappings'][:3]):  # 只显示前3个
            print(f"  文件 {i+1}:")
            print(f"    嵌入路径: {mapping['embedded_path']}")
            print(f"    当前名称: {mapping['current_name']}")
            print(f"    检测类型: {mapping['detected_type']}")
            print(f"    原始名称: {mapping['original_name']} (需要用户填写)")
        
        if len(mapping_data['mappings']) > 3:
            print(f"    ... 还有 {len(mapping_data['mappings']) - 3} 个文件")
        
        # 2. 创建一个示例映射文件（模拟用户填写）
        print("\n步骤2: 创建示例映射文件（模拟用户填写原始文件名）")
        
        # 模拟用户填写的原始文件名
        sample_names = [
            "课程设计文档.docx",
            "实验手册.docx", 
            "评分标准.xlsx",
            "课程大纲.pptx",
            "作业模板.docx",
            "考试题库.xlsx",
            "教学计划.docx",
            "学习资料.xlsx",
            "课程介绍.pptx",
            "参考资料.docx"
        ]
        
        # 填写示例原始文件名
        for i, mapping in enumerate(mapping_data['mappings']):
            if i < len(sample_names):
                mapping['original_name'] = sample_names[i]
                mapping['description'] = f"这是第{i+1}个嵌入的文档文件"
        
        # 保存示例映射文件
        sample_mapping_file = f"{Path(ppt_file).stem}_sample_mapping.json"
        with open(sample_mapping_file, 'w', encoding='utf-8') as f:
            json.dump(mapping_data, f, ensure_ascii=False, indent=2)
        
        print(f"示例映射文件已创建: {sample_mapping_file}")
        
        # 3. 使用示例映射文件提取
        print("\n步骤3: 使用示例映射文件提取嵌入对象")
        extractor_with_mapping = EnhancedPPTExtractor("extracted_objects_with_mapping")
        success = extractor_with_mapping.extract_with_mapping(ppt_file, sample_mapping_file)
        
        if success:
            print("\n✓ 使用映射文件提取成功！")
        else:
            print("\n✗ 使用映射文件提取失败")
    
    # 4. 演示智能命名提取
    print("\n步骤4: 演示智能命名提取")
    smart_extractor = EnhancedPPTExtractor("extracted_objects_smart_naming")
    success = smart_extractor.extract_with_smart_naming(ppt_file)
    
    if success:
        print("\n✓ 智能命名提取成功！")
    else:
        print("\n✗ 智能命名提取失败")
    
    # 5. 总结
    print("\n=== 演示总结 ===")
    print("\n增强版提取器提供了三种解决方案:")
    print("\n1. 文件名映射模板（推荐）:")
    print("   - 创建JSON模板文件")
    print("   - 用户手动填写原始中文文件名")
    print("   - 提取时使用用户指定的文件名")
    print("   - 完全保留原始文件名")
    
    print("\n2. 智能命名策略:")
    print("   - 根据文件类型自动生成中文描述性名称")
    print("   - 如: Word文档_01.docx, Excel表格_01.xlsx")
    print("   - 无需用户干预，但不是原始文件名")
    
    print("\n3. 使用建议:")
    print("   - 如果需要保留原始文件名，使用方案1")
    print("   - 如果只需要有意义的文件名，使用方案2")
    print("   - 映射模板可以重复使用和修改")
    
    print("\n文件输出位置:")
    print(f"   - 映射模板: {mapping_file if mapping_file else '未创建'}")
    print(f"   - 示例映射: {sample_mapping_file if 'sample_mapping_file' in locals() else '未创建'}")
    print("   - 映射提取结果: extracted_objects_with_mapping/")
    print("   - 智能命名结果: extracted_objects_smart_naming/")

if __name__ == "__main__":
    demo_enhanced_extractor()