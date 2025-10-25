@echo off

REM 运行Word图片提取工具并记录日志

set LOG_FILE=run_log.txt

echo 开始运行Word图片提取工具... > %LOG_FILE%
echo 当前时间: %date% %time% >> %LOG_FILE%

echo 请确保已安装所需依赖: >> %LOG_FILE%
echo - python-docx >> %LOG_FILE%
echo - pillow >> %LOG_FILE%

echo. >> %LOG_FILE%
echo 正在运行脚本... >> %LOG_FILE%

REM 运行Python脚本并捕获输出和错误
python word_image_extractor.py >> %LOG_FILE% 2>&1

REM 检查退出码
if %ERRORLEVEL% neq 0 (
    echo. >> %LOG_FILE%
    echo 程序出现错误，错误码: %ERRORLEVEL% >> %LOG_FILE%
    echo 请查看日志文件获取详细信息 >> %LOG_FILE%
    echo 错误已记录到 %LOG_FILE%，请查看详细信息
    pause
    exit /b %ERRORLEVEL%
) else (
    echo. >> %LOG_FILE%
    echo 程序成功完成 >> %LOG_FILE%
    echo 程序已成功运行，日志已保存到 %LOG_FILE%
)

echo. >> %LOG_FILE%
echo 结束时间: %date% %time% >> %LOG_FILE%

echo. 
echo 按任意键查看日志...
type %LOG_FILE%
pause