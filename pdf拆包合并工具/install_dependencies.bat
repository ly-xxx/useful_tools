@echo off
chcp 65001 > nul
echo 文档处理工具 - 依赖项安装(conda版)
echo =======================================

:: 检查conda是否安装
where conda >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo 错误：未找到conda。请安装Anaconda或Miniconda。
    pause
    exit /b 1
)

echo 创建conda环境(doc_processor)并安装所需依赖项...
echo 这可能需要几分钟时间...

:: 创建conda环境
conda create -y -n doc_processor python=3.8

:: 激活环境
call conda activate doc_processor

:: 安装依赖
pip install patool pyunpack python-docx docx2pdf PyPDF2 pikepdf pymupdf python-pptx pywin32 openpyxl

if %ERRORLEVEL% neq 0 (
    echo 安装依赖项时出错。请检查网络连接并重试。
    pause
    exit /b 1
)

echo.
echo 所有依赖项安装完成！
echo.
echo 使用方法:
echo 1. 激活环境: conda activate doc_processor
echo 2. 运行程序: python doc_processor.py [可选:文件夹路径]
echo.
pause 