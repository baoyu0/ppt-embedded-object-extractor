import zipfile
import xml.etree.ElementTree as ET
import re

ppt_file = r'd:\00-深圳华云\13-自服务课程开发\大语言模型\程燕霞\【请查收评审建议+进度+提交PDF版】开发者人才培养华云伙伴：《大语言模型》PPT_讲义实验手册_代码评审结果+交付件进度+PDF版\1\课程共建交付件清单和开发顺序0828 - 20250903145602.pptx'

with zipfile.ZipFile(ppt_file, 'r') as zip_ref:
    print("=== 搜索所有XML文件中的中文内容 ===")
    
    # 获取所有XML文件
    xml_files = [f for f in zip_ref.namelist() if f.endswith('.xml')]
    
    for xml_file in xml_files:
        print(f'\n检查文件: {xml_file}')
        try:
            content = zip_ref.read(xml_file).decode('utf-8')
            
            # 查找包含中文字符的内容
            chinese_matches = re.findall(r'[\u4e00-\u9fff][^<>"]*[\u4e00-\u9fff]*', content)
            
            if chinese_matches:
                print(f'  发现中文内容:')
                for match in chinese_matches[:10]:  # 只显示前10个
                    clean_match = match.strip()
                    if len(clean_match) >= 2:
                        print(f'    {clean_match}')
            
            # 特别查找name属性中的中文
            name_matches = re.findall(r'name="([^"]*[\u4e00-\u9fff][^"]*?)"', content)
            if name_matches:
                print(f'  name属性中的中文:')
                for match in name_matches:
                    print(f'    {match}')
                    
            # 查找title属性中的中文
            title_matches = re.findall(r'title="([^"]*[\u4e00-\u9fff][^"]*?)"', content)
            if title_matches:
                print(f'  title属性中的中文:')
                for match in title_matches:
                    print(f'    {match}')
                    
            # 查找文本节点中的中文
            text_matches = re.findall(r'>([^<]*[\u4e00-\u9fff][^<]*?)<', content)
            if text_matches:
                print(f'  文本节点中的中文:')
                for match in text_matches[:5]:  # 只显示前5个
                    clean_match = match.strip()
                    if len(clean_match) >= 2:
                        print(f'    {clean_match}')
                        
        except Exception as e:
            print(f'  处理文件时出错: {e}')
    
    print("\n=== 特别检查docProps文件 ===")
    docprops_files = [f for f in zip_ref.namelist() if 'docProps' in f]
    for file in docprops_files:
        print(f'\n检查: {file}')
        try:
            content = zip_ref.read(file).decode('utf-8')
            print(f'内容预览: {content[:500]}...')
            
            # 查找中文内容
            chinese_matches = re.findall(r'[\u4e00-\u9fff][^<>"]*', content)
            if chinese_matches:
                print('发现中文:')
                for match in chinese_matches:
                    print(f'  {match}')
        except Exception as e:
            print(f'处理时出错: {e}')