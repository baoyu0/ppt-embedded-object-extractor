@echo off
echo 正在创建Python虚拟环境...
python -m venv ppt_extract_env
echo 虚拟环境创建完成！

echo 正在激活虚拟环境...
call ppt_extract_env\Scripts\activate
echo 虚拟环境已激活！

echo 正在安装依赖包...
pip install --upgrade pip
pip install -r requirements.txt
echo 依赖包安装完成！

echo.
echo 环境设置完成！
echo 要激活虚拟环境，请运行: ppt_extract_env\Scripts\activate
echo 要运行程序，请使用: python ppt_extractor.py
pause