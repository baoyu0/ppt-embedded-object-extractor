# PPT嵌入对象提取工具

一个强大的Python工具，用于从PowerPoint文件中提取所有嵌入的Word、Excel和其他文件对象，并将它们完整保存到指定文件夹中。

## 🌟 功能特性

- ✅ **多格式支持**: 支持PPTX和PPT格式文件
- 📁 **智能识别**: 自动识别和分类不同类型的嵌入文件（Word、Excel、PDF等）
- 🔍 **文件类型检测**: 基于文件头签名的精确文件类型识别
- 📊 **详细报告**: 提供完整的提取报告和统计信息
- 🛡️ **错误处理**: 完善的异常处理和错误报告机制
- 📝 **日志记录**: 可选的详细日志记录功能
- 🖥️ **多种模式**: 支持命令行和交互式两种使用模式
- 🔄 **文件完整性**: 保持原始文件名和内容完整性

## 📋 系统要求

- Python 3.7 或更高版本
- Windows、macOS 或 Linux 操作系统
- 至少 100MB 可用磁盘空间

## 🚀 快速开始

### 1. 环境设置

#### Windows 用户
```bash
# 运行自动化设置脚本
setup_env.bat
```

#### Linux/macOS 用户
```bash
# 运行自动化设置脚本
chmod +x setup_env.sh
./setup_env.sh
```

#### 手动设置
```bash
# 创建虚拟环境
python -m venv ppt_extract_env

# 激活虚拟环境
# Windows:
ppt_extract_env\Scripts\activate
# Linux/macOS:
source ppt_extract_env/bin/activate

# 安装依赖
pip install -r requirements.txt
```

### 2. 基本使用

#### 命令行模式
```bash
# 基本用法
python main.py presentation.pptx output_folder

# 使用命名参数
python main.py -i presentation.pptx -o output_folder

# 启用日志记录
python main.py presentation.pptx output_folder --log extract.log

# 禁用控制台输出
python main.py presentation.pptx output_folder --no-console-log
```

#### 交互模式
```bash
# 启动交互模式
python main.py --interactive
```

## 📖 详细使用说明

### 命令行参数

| 参数 | 描述 | 示例 |
|------|------|------|
| `input_file` | 输入的PPT文件路径 | `presentation.pptx` |
| `output_dir` | 输出目录路径 | `extracted_files` |
| `-i, --input` | 输入文件路径（替代位置参数） | `-i presentation.pptx` |
| `-o, --output` | 输出目录路径（替代位置参数） | `-o extracted_files` |
| `--log` | 日志文件路径 | `--log extract.log` |
| `--no-console-log` | 禁用控制台日志输出 | `--no-console-log` |
| `--interactive` | 启用交互模式 | `--interactive` |
| `--version` | 显示版本信息 | `--version` |

### 支持的文件类型

| 文件类型 | 扩展名 | 描述 |
|----------|--------|------|
| Word文档 | `.docx`, `.doc` | Microsoft Word文档 |
| Excel表格 | `.xlsx`, `.xls` | Microsoft Excel工作簿 |
| PowerPoint | `.pptx`, `.ppt` | Microsoft PowerPoint演示文稿 |
| PDF文档 | `.pdf` | Adobe PDF文档 |
| 文本文件 | `.txt`, `.rtf` | 纯文本和富文本文件 |
| 图像文件 | `.png`, `.jpg`, `.gif` | 常见图像格式 |
| 其他格式 | 各种 | 根据文件头自动识别 |

### 输出结果

提取完成后，工具会生成：

1. **提取的文件**: 保存在指定输出目录中，按原始文件名命名
2. **详细报告**: 显示提取统计信息和文件列表
3. **日志文件**: （可选）包含详细的操作日志
4. **错误报告**: 列出任何提取失败的文件及原因

### 示例输出

```
============================================================
PPT嵌入对象提取报告
============================================================
总计发现文件: 5
成功提取: 4
提取失败: 1
处理时间: 2.34 秒
总提取大小: 2.5 MB

✅ 成功提取的文件:
----------------------------------------

📁 Office文档 (3 个文件):
  ✓ 销售报告.xlsx
    类型: Excel工作簿
    大小: 1.2 MB
    来源: 幻灯片 2

  ✓ 产品介绍.docx
    类型: Word文档
    大小: 856 KB
    来源: 幻灯片 3

  ✓ 附件演示.pptx
    类型: PowerPoint演示文稿
    大小: 445 KB
    来源: 幻灯片 5

📁 PDF文档 (1 个文件):
  ✓ 技术规格.pdf
    类型: PDF文档
    大小: 234 KB
    来源: 幻灯片 4

❌ 提取失败的文件:
----------------------------------------
  ✗ 损坏文件.unknown
    错误类型: PPTExtractorError
    错误信息: 文件格式不受支持
    来源: 幻灯片 6

============================================================
```

## 🔧 高级功能

### 编程接口

```python
from ppt_extractor import PPTExtractor

# 创建提取器实例
extractor = PPTExtractor(log_file='extract.log', enable_console_log=True)

# 执行提取
result = extractor.extract_embedded_objects('presentation.pptx', 'output_folder')

# 显示报告
extractor.print_extraction_report(result)

# 访问结果数据
print(f"成功提取 {len(result['success'])} 个文件")
print(f"失败 {len(result['failed'])} 个文件")
```

### 自定义文件类型检测

```python
from file_type_detector import FileTypeDetector

# 创建文件类型检测器
detector = FileTypeDetector()

# 检测文件类型
with open('unknown_file.bin', 'rb') as f:
    file_data = f.read()
    
file_ext, file_type, mime_type = detector.detect_file_type(file_data, 'unknown_file.bin')
print(f"文件类型: {file_type}")
print(f"MIME类型: {mime_type}")
```

### 错误处理

```python
from error_handler import ErrorHandler, safe_execute

# 创建错误处理器
error_handler = ErrorHandler('error.log')

# 安全执行操作
def risky_operation():
    # 可能出错的代码
    pass

result = safe_execute(
    risky_operation,
    error_message="操作失败",
    context={'operation': 'test'}
)
```

## 🐛 故障排除

### 常见问题

**Q: 提示"文件不存在"错误**
A: 请检查文件路径是否正确，确保文件存在且可访问。

**Q: 提示"权限不足"错误**
A: 请确保对输入文件有读取权限，对输出目录有写入权限。

**Q: 某些文件提取失败**
A: 可能是文件损坏或格式不受支持，查看详细错误信息进行诊断。

**Q: 程序运行缓慢**
A: 大型PPT文件可能需要较长时间处理，请耐心等待。

### 调试模式

启用详细日志记录：
```bash
python main.py presentation.pptx output_folder --log debug.log
```

查看日志文件获取详细的错误信息和处理过程。

## 📁 项目结构

```
ppt-extractor/
├── main.py                 # 主程序入口
├── ppt_extractor.py        # 核心提取功能
├── file_type_detector.py   # 文件类型检测
├── error_handler.py        # 错误处理机制
├── requirements.txt        # Python依赖
├── setup_env.bat          # Windows环境设置
├── setup_env.sh           # Linux/macOS环境设置
└── README.md              # 使用说明
```

## 🤝 贡献指南

欢迎提交问题报告和功能请求！如果您想贡献代码：

1. Fork 本项目
2. 创建功能分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 开启 Pull Request

## 📄 许可证

本项目采用 MIT 许可证 - 查看 [LICENSE](LICENSE) 文件了解详情。

## 🙏 致谢

- [python-pptx](https://python-pptx.readthedocs.io/) - PowerPoint文件处理
- [Pillow](https://pillow.readthedocs.io/) - 图像处理支持
- [tqdm](https://tqdm.github.io/) - 进度条显示

## 📞 支持

如果您遇到问题或需要帮助，请：

1. 查看本文档的故障排除部分
2. 检查 [Issues](../../issues) 页面是否有类似问题
3. 创建新的 Issue 描述您的问题

---

**版本**: 1.0.0  
**最后更新**: 2024年1月