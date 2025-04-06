# 文档处理工具

这个工具可以解压所有压缩包（包括嵌套压缩包），将Word和PDF文件转换并合并为一个PDF文件。

## 安装方法 (两种方式)

### 方式一：使用批处理脚本一键安装

运行 `install_dependencies.bat` 自动创建conda环境并安装所有依赖项。

### 方式二：使用conda环境文件

```bash
# 创建conda环境
conda env create -f environment.yml

# 激活环境
conda activate doc_processor
```

## 使用方法

1. 激活conda环境：
```bash
conda activate doc_processor
```

2. 处理当前目录下的文件：
```bash
python doc_processor.py
```

3. 处理指定目录下的文件：
```bash
python doc_processor.py D:\文档路径
```

## 功能

- 自动解压缩各种格式（zip、rar、7z等）
- 支持嵌套压缩包的解压
- 将Word文档和PPT转换为PDF
- 合并所有PDF为一个文件
- 输出文件保存在原始目录下

## 系统要求

- Windows 系统
- Anaconda或Miniconda
- Microsoft Office (用于Word和PowerPoint转换) 