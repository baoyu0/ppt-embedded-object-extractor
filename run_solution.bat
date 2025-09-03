
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
