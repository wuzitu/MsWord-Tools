@echo off
REM Word表格图片提取工具(GPT-word.py)打包脚本

setlocal

echo ===== Word表格图片提取工具打包开始 =====

REM 检查Python是否安装
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo 错误：未找到Python，请先安装Python
    pause
    exit /b 1
)

REM 升级pip
echo 升级pip...
python -m pip install --upgrade pip

REM 安装打包工具和依赖
echo 安装打包工具和项目依赖...
python -m pip install pyinstaller
python -m pip install -r requirements.txt

REM 执行打包
echo 开始打包可执行文件...
pyinstaller --onefile --windowed --name="Word表格图片提取工具" --icon=nul --add-data="requirements.txt;." GPT-word.py

if %errorlevel% neq 0 (
    echo 打包失败！
    pause
    exit /b 1
)

echo 打包成功！可执行文件位于 dist 目录中
echo 可执行文件：dist\Word表格图片提取工具.exe
echo ===== 打包完成 =====

pause