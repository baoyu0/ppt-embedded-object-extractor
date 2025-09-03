#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PPT嵌入对象提取工具 - 使用示例

本文件展示了如何使用PPT嵌入对象提取工具的各种功能。
"""

import os
import sys
from pathlib import Path

# 添加当前目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from ppt_extractor import PPTExtractor
from file_type_detector import FileTypeDetector
from error_handler import ErrorHandler, safe_execute, PPTExtractorError


def example_basic_extraction():
    """
    示例1: 基本的文件提取
    """
    print("\n" + "="*50)
    print("示例1: 基本的文件提取")
    print("="*50)
    
    # 创建提取器
    extractor = PPTExtractor()
    
    # 假设的PPT文件路径
    ppt_file = "sample_presentation.pptx"
    output_dir = "extracted_files"
    
    # 检查文件是否存在
    if not os.path.exists(ppt_file):
        print(f"⚠️  示例文件 {ppt_file} 不存在，跳过此示例")
        return
    
    try:
        # 执行提取
        result = extractor.extract_embedded_objects(ppt_file, output_dir)
        
        # 显示报告
        extractor.print_extraction_report(result)
        
    except Exception as e:
        print(f"❌ 提取失败: {e}")


def example_with_logging():
    """
    示例2: 带日志记录的提取
    """
    print("\n" + "="*50)
    print("示例2: 带日志记录的提取")
    print("="*50)
    
    # 创建带日志的提取器
    log_file = "extraction.log"
    extractor = PPTExtractor(log_file=log_file, enable_console_log=True)
    
    ppt_file = "sample_presentation.pptx"
    output_dir = "extracted_with_log"
    
    if not os.path.exists(ppt_file):
        print(f"⚠️  示例文件 {ppt_file} 不存在，跳过此示例")
        return
    
    try:
        result = extractor.extract_embedded_objects(ppt_file, output_dir)
        print(f"\n📝 日志已保存到: {log_file}")
        
        # 显示日志内容的前几行
        if os.path.exists(log_file):
            with open(log_file, 'r', encoding='utf-8') as f:
                lines = f.readlines()[-10:]  # 显示最后10行
                print("\n📋 日志内容（最后10行）:")
                print("-" * 40)
                for line in lines:
                    print(line.strip())
        
    except Exception as e:
        print(f"❌ 提取失败: {e}")


def example_file_type_detection():
    """
    示例3: 文件类型检测功能
    """
    print("\n" + "="*50)
    print("示例3: 文件类型检测功能")
    print("="*50)
    
    # 创建文件类型检测器
    detector = FileTypeDetector()
    
    # 测试不同的文件头
    test_data = [
        (b'PK\x03\x04', 'test.docx', 'Word文档测试'),
        (b'%PDF-1.4', 'test.pdf', 'PDF文档测试'),
        (b'\xd0\xcf\x11\xe0', 'test.doc', '旧版Office文档测试'),
        (b'\xff\xd8\xff\xe0', 'test.jpg', 'JPEG图像测试'),
        (b'\x89PNG\r\n\x1a\n', 'test.png', 'PNG图像测试'),
    ]
    
    print("🔍 文件类型检测结果:")
    print("-" * 40)
    
    for file_data, filename, description in test_data:
        try:
            # 扩展数据以模拟真实文件
            extended_data = file_data + b'\x00' * 100
            
            file_ext, file_type, mime_type = detector.detect_file_type(extended_data, filename)
            category = detector.get_file_category(file_ext[1:] if file_ext.startswith('.') else file_ext)
            
            print(f"📄 {description}:")
            print(f"   扩展名: {file_ext}")
            print(f"   类型: {file_type}")
            print(f"   MIME: {mime_type}")
            print(f"   分类: {category}")
            print()
            
        except Exception as e:
            print(f"❌ 检测失败 ({description}): {e}")


def example_error_handling():
    """
    示例4: 错误处理机制
    """
    print("\n" + "="*50)
    print("示例4: 错误处理机制")
    print("="*50)
    
    # 创建错误处理器
    error_handler = ErrorHandler("error_example.log")
    
    # 模拟各种错误情况
    def test_file_not_found():
        raise FileNotFoundError("测试文件不存在")
    
    def test_permission_error():
        raise PermissionError("测试权限不足")
    
    def test_custom_error():
        raise PPTExtractorError("测试自定义错误", None)
    
    # 测试安全执行
    test_functions = [
        (test_file_not_found, "文件不存在测试"),
        (test_permission_error, "权限错误测试"),
        (test_custom_error, "自定义错误测试"),
    ]
    
    print("🛡️ 错误处理测试:")
    print("-" * 40)
    
    for test_func, description in test_functions:
        try:
            result = safe_execute(
                test_func,
                error_message=f"{description}失败",
                context={'test': description}
            )
            print(f"✅ {description}: 意外成功")
        except Exception as e:
            print(f"❌ {description}: {type(e).__name__} - {e}")
    
    # 显示错误摘要
    if hasattr(error_handler, 'get_error_summary'):
        try:
            summary = error_handler.get_error_summary()
            if summary:
                print("\n📊 错误摘要:")
                for error_type, count in summary.items():
                    print(f"   {error_type}: {count} 次")
        except:
            print("\n📊 错误摘要功能不可用")


def example_batch_processing():
    """
    示例5: 批量处理多个PPT文件
    """
    print("\n" + "="*50)
    print("示例5: 批量处理多个PPT文件")
    print("="*50)
    
    # 查找当前目录下的所有PPT文件
    ppt_files = []
    for ext in ['*.ppt', '*.pptx']:
        ppt_files.extend(Path('.').glob(ext))
    
    if not ppt_files:
        print("⚠️  当前目录下没有找到PPT文件，创建示例文件列表")
        ppt_files = ['presentation1.pptx', 'presentation2.pptx', 'old_format.ppt']
    
    print(f"📁 发现 {len(ppt_files)} 个PPT文件")
    
    # 创建提取器
    extractor = PPTExtractor(log_file="batch_extraction.log")
    
    # 批量处理
    total_extracted = 0
    total_failed = 0
    
    for i, ppt_file in enumerate(ppt_files, 1):
        print(f"\n🔄 处理文件 {i}/{len(ppt_files)}: {ppt_file}")
        
        if not os.path.exists(str(ppt_file)):
            print(f"   ⚠️  文件不存在，跳过")
            continue
        
        try:
            output_dir = f"batch_output/file_{i}"
            result = extractor.extract_embedded_objects(str(ppt_file), output_dir)
            
            summary = result.get('summary', {})
            extracted = summary.get('extracted', 0)
            failed = summary.get('failed', 0)
            
            total_extracted += extracted
            total_failed += failed
            
            print(f"   ✅ 提取成功: {extracted} 个文件")
            if failed > 0:
                print(f"   ❌ 提取失败: {failed} 个文件")
                
        except Exception as e:
            print(f"   💥 处理失败: {e}")
            total_failed += 1
    
    print(f"\n📊 批量处理完成:")
    print(f"   总计提取成功: {total_extracted} 个文件")
    print(f"   总计提取失败: {total_failed} 个文件")


def main():
    """
    运行所有示例
    """
    print("🚀 PPT嵌入对象提取工具 - 使用示例")
    print("=" * 60)
    
    # 运行各个示例
    examples = [
        example_basic_extraction,
        example_with_logging,
        example_file_type_detection,
        example_error_handling,
        example_batch_processing,
    ]
    
    for example_func in examples:
        try:
            example_func()
        except KeyboardInterrupt:
            print("\n⏹️  用户中断操作")
            break
        except Exception as e:
            print(f"\n💥 示例执行失败: {e}")
    
    print("\n🎉 所有示例执行完成！")
    print("\n💡 提示:")
    print("   - 要运行真实的提取操作，请准备一些包含嵌入对象的PPT文件")
    print("   - 查看生成的日志文件了解详细的处理过程")
    print("   - 使用 'python main.py --help' 查看命令行选项")


if __name__ == '__main__':
    main()