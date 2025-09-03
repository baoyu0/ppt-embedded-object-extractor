#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PPT嵌入对象提取工具 - 主程序入口

这个工具可以从PowerPoint文件中提取所有嵌入的Word、Excel和其他文件对象，
并将它们完整保存到指定文件夹中。

作者: AI Assistant
版本: 1.0.0
"""

import os
import sys
import argparse
from pathlib import Path
from typing import Optional, Tuple

# 添加当前目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from ppt_extractor import PPTExtractor
from error_handler import PPTExtractorError, ErrorCode, ErrorHandler, safe_execute


def create_argument_parser() -> argparse.ArgumentParser:
    """
    创建命令行参数解析器
    """
    parser = argparse.ArgumentParser(
        description='从PPT文件中提取嵌入的Word、Excel和其他文件对象',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例用法:
  %(prog)s presentation.pptx output_folder
  %(prog)s -i presentation.pptx -o output_folder --log extract.log
  %(prog)s --interactive
        """
    )
    
    # 位置参数
    parser.add_argument(
        'input_file',
        nargs='?',
        help='输入的PPT文件路径 (.ppt 或 .pptx)'
    )
    
    parser.add_argument(
        'output_dir',
        nargs='?',
        help='输出目录路径（提取的文件将保存在此目录中）'
    )
    
    # 可选参数
    parser.add_argument(
        '-i', '--input',
        dest='input_file_alt',
        help='输入的PPT文件路径（替代位置参数）'
    )
    
    parser.add_argument(
        '-o', '--output',
        dest='output_dir_alt',
        help='输出目录路径（替代位置参数）'
    )
    
    parser.add_argument(
        '--log',
        dest='log_file',
        help='日志文件路径（可选）'
    )
    
    parser.add_argument(
        '--no-console-log',
        action='store_true',
        help='禁用控制台日志输出'
    )
    
    parser.add_argument(
        '--interactive',
        action='store_true',
        help='启用交互模式'
    )
    
    parser.add_argument(
        '--version',
        action='version',
        version='PPT嵌入对象提取工具 v1.0.0'
    )
    
    return parser


def interactive_mode() -> Tuple[str, str, Optional[str]]:
    """
    交互模式：引导用户输入参数
    
    Returns:
        (input_file, output_dir, log_file)
    """
    print("\n" + "="*60)
    print("PPT嵌入对象提取工具 - 交互模式")
    print("="*60)
    
    # 获取输入文件
    while True:
        input_file = input("\n请输入PPT文件路径: ").strip().strip('"\'')
        if not input_file:
            print("❌ 请输入有效的文件路径")
            continue
        
        if not os.path.exists(input_file):
            print(f"❌ 文件不存在: {input_file}")
            continue
        
        if not input_file.lower().endswith(('.ppt', '.pptx')):
            print("❌ 请选择PPT或PPTX格式的文件")
            continue
        
        break
    
    # 获取输出目录
    while True:
        output_dir = input("\n请输入输出目录路径: ").strip().strip('"\'')
        if not output_dir:
            print("❌ 请输入有效的目录路径")
            continue
        
        # 如果目录不存在，询问是否创建
        if not os.path.exists(output_dir):
            create = input(f"目录 '{output_dir}' 不存在，是否创建？ (y/n): ").strip().lower()
            if create in ['y', 'yes', '是']:
                try:
                    os.makedirs(output_dir, exist_ok=True)
                    print(f"✅ 已创建目录: {output_dir}")
                    break
                except Exception as e:
                    print(f"❌ 创建目录失败: {e}")
                    continue
            else:
                continue
        else:
            break
    
    # 获取日志文件（可选）
    log_file = input("\n请输入日志文件路径（可选，直接回车跳过）: ").strip().strip('"\'')
    if not log_file:
        log_file = None
    
    return input_file, output_dir, log_file


def validate_arguments(args) -> Tuple[str, str, Optional[str]]:
    """
    验证和处理命令行参数
    
    Returns:
        (input_file, output_dir, log_file)
    """
    # 确定输入文件
    input_file = args.input_file or args.input_file_alt
    if not input_file:
        raise PPTExtractorError(
            "请指定输入的PPT文件路径",
            ErrorCode.INVALID_ARGUMENTS
        )
    
    # 确定输出目录
    output_dir = args.output_dir or args.output_dir_alt
    if not output_dir:
        raise PPTExtractorError(
            "请指定输出目录路径",
            ErrorCode.INVALID_ARGUMENTS
        )
    
    # 验证输入文件
    if not os.path.exists(input_file):
        raise PPTExtractorError(
            f"输入文件不存在: {input_file}",
            ErrorCode.FILE_NOT_FOUND
        )
    
    if not input_file.lower().endswith(('.ppt', '.pptx')):
        raise PPTExtractorError(
            "输入文件必须是PPT或PPTX格式",
            ErrorCode.FILE_FORMAT_UNSUPPORTED
        )
    
    return input_file, output_dir, args.log_file


def main():
    """
    主函数
    """
    parser = create_argument_parser()
    args = parser.parse_args()
    
    try:
        # 处理参数
        if args.interactive:
            input_file, output_dir, log_file = interactive_mode()
            enable_console_log = True
        else:
            input_file, output_dir, log_file = validate_arguments(args)
            enable_console_log = not args.no_console_log
        
        # 显示开始信息
        print("\n" + "="*60)
        print("PPT嵌入对象提取工具")
        print("="*60)
        print(f"输入文件: {input_file}")
        print(f"输出目录: {output_dir}")
        if log_file:
            print(f"日志文件: {log_file}")
        print("-" * 60)
        
        # 创建错误处理器和提取器
        error_handler = ErrorHandler(log_file=log_file, enable_console=enable_console_log)
        extractor = PPTExtractor(log_file=log_file, enable_console_log=enable_console_log)
        
        def extract_operation():
            return extractor.extract_embedded_objects(input_file, output_dir)
        
        # 安全执行提取操作
        result = safe_execute(
            extract_operation,
            error_handler=error_handler
        )
        
        # 显示结果报告
        extractor.print_extraction_report(result)
        
        # 根据结果设置退出码
        summary = result.get('summary', {})
        if summary.get('failed', 0) > 0:
            print("\n⚠️  部分文件提取失败，请查看上述错误信息")
            sys.exit(1)
        elif summary.get('extracted', 0) == 0:
            print("\n📝 未发现任何嵌入对象")
            sys.exit(0)
        else:
            print("\n🎉 所有嵌入对象提取成功！")
            sys.exit(0)
    
    except KeyboardInterrupt:
        print("\n\n⏹️  用户中断操作")
        sys.exit(130)
    
    except PPTExtractorError as e:
        print(f"\n❌ 错误: {e}")
        if e.error_code:
            print(f"错误代码: {e.error_code.value}")
        sys.exit(1)
    
    except Exception as e:
        print(f"\n💥 未预期的错误: {e}")
        print("请检查输入参数或联系技术支持")
        sys.exit(1)


if __name__ == '__main__':
    main()