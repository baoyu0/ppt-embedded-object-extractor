import zipfile
import struct
import os
import re

ppt_file = r'd:\00-深圳华云\13-自服务课程开发\大语言模型\程燕霞\【请查收评审建议+进度+提交PDF版】开发者人才培养华云伙伴：《大语言模型》PPT_讲义实验手册_代码评审结果+交付件进度+PDF版\1\课程共建交付件清单和开发顺序0828 - 20250903145602.pptx'

with zipfile.ZipFile(ppt_file, 'r') as zip_ref:
    # 查找嵌入对象文件
    embedding_files = [f for f in zip_ref.namelist() if f.startswith('ppt/embeddings/')]
    
    print(f'找到 {len(embedding_files)} 个嵌入文件:')
    for i, file_path in enumerate(embedding_files[:5]):  # 只检查前5个
        print(f'\n=== 分析 {file_path} ===')
        try:
            data = zip_ref.read(file_path)
            print(f'文件大小: {len(data)} 字节')
            
            # 检查文件头
            if len(data) >= 8:
                header = data[:8]
                print(f'文件头 (hex): {header.hex()}')
                print(f'文件头 (ascii): {"".join([chr(b) if 32 <= b <= 126 else "." for b in header])}')
            
            # 查找可能的中文字符串
            try:
                text_data = data.decode('utf-8', errors='ignore')
                chinese_chars = []
                for char in text_data:
                    if '\u4e00' <= char <= '\u9fff':  # 中文字符范围
                        chinese_chars.append(char)
                
                if chinese_chars:
                    print(f'发现中文字符: {"".join(chinese_chars[:50])}')  # 只显示前50个字符
                    
                    # 尝试找到包含中文的完整字符串
                    chinese_strings = re.findall(r'[\u4e00-\u9fff][\u4e00-\u9fff\w\s]*', text_data)
                    if chinese_strings:
                        print('包含中文的字符串:')
                        for s in chinese_strings[:10]:  # 只显示前10个
                            print(f'  {s}')
                else:
                    print('未发现中文字符')
            except:
                print('UTF-8解码失败，尝试其他编码')
                
                # 尝试GBK编码
                try:
                    text_data = data.decode('gbk', errors='ignore')
                    chinese_chars = []
                    for char in text_data:
                        if '\u4e00' <= char <= '\u9fff':  # 中文字符范围
                            chinese_chars.append(char)
                    
                    if chinese_chars:
                        print(f'GBK编码发现中文字符: {"".join(chinese_chars[:50])}')
                        chinese_strings = re.findall(r'[\u4e00-\u9fff][\u4e00-\u9fff\w\s]*', text_data)
                        if chinese_strings:
                            print('GBK编码包含中文的字符串:')
                            for s in chinese_strings[:10]:
                                print(f'  {s}')
                except:
                    print('GBK解码也失败')
                
        except Exception as e:
            print(f'分析 {file_path} 时出错: {e}')