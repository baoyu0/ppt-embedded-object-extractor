import zipfile
import struct
import os

def analyze_ole_compound_doc(data):
    """分析OLE复合文档结构"""
    try:
        # OLE复合文档的签名
        ole_signature = b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'
        if not data.startswith(ole_signature):
            return None
            
        print("这是一个OLE复合文档")
        
        # 尝试查找可能的文件名信息
        # 在OLE文档中，文件名通常以UTF-16编码存储
        
        # 查找可能的UTF-16编码的中文字符串
        for i in range(0, len(data) - 4, 2):
            try:
                # 尝试解码UTF-16LE
                chunk = data[i:i+40]  # 取40字节
                if len(chunk) >= 4:
                    text = chunk.decode('utf-16le', errors='ignore')
                    # 检查是否包含中文字符
                    if any('\u4e00' <= char <= '\u9fff' for char in text):
                        # 清理字符串，只保留可打印字符
                        clean_text = ''.join(char for char in text if char.isprintable())
                        if len(clean_text) >= 2 and any('\u4e00' <= char <= '\u9fff' for char in clean_text):
                            print(f"可能的中文文件名 (UTF-16LE): {clean_text}")
            except:
                continue
                
        # 查找可能的UTF-8编码的中文字符串
        text_utf8 = data.decode('utf-8', errors='ignore')
        import re
        chinese_patterns = re.findall(r'[\u4e00-\u9fff][\u4e00-\u9fff\w\s\.]*', text_utf8)
        for pattern in chinese_patterns:
            if len(pattern.strip()) >= 2:
                print(f"可能的中文文件名 (UTF-8): {pattern.strip()}")
                
    except Exception as e:
        print(f"分析OLE文档时出错: {e}")
        return None

ppt_file = r'd:\00-深圳华云\13-自服务课程开发\大语言模型\程燕霞\【请查收评审建议+进度+提交PDF版】开发者人才培养华云伙伴：《大语言模型》PPT_讲义实验手册_代码评审结果+交付件进度+PDF版\1\课程共建交付件清单和开发顺序0828 - 20250903145602.pptx'

with zipfile.ZipFile(ppt_file, 'r') as zip_ref:
    # 查找.bin文件（这些通常是OLE对象）
    ole_files = [f for f in zip_ref.namelist() if f.endswith('.bin') and 'embeddings' in f]
    
    print(f'找到 {len(ole_files)} 个OLE对象文件:')
    for i, file_path in enumerate(ole_files[:3]):  # 只检查前3个
        print(f'\n=== 分析 {file_path} ===')
        try:
            data = zip_ref.read(file_path)
            print(f'文件大小: {len(data)} 字节')
            
            # 检查文件头
            if len(data) >= 8:
                header = data[:8]
                print(f'文件头 (hex): {header.hex()}')
            
            analyze_ole_compound_doc(data)
                
        except Exception as e:
            print(f'分析 {file_path} 时出错: {e}')