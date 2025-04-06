#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
文档处理工具
功能：解压所有压缩包（包括嵌套压缩包），将Word和PDF文件转换并合并为一个大PDF
"""

import os
import sys
import shutil
import tempfile
import traceback
from pathlib import Path
import subprocess

# 用于提取压缩文件
import patoolib
from pyunpack import Archive, PatoolError

# 用于PDF处理
import PyPDF2
from pikepdf import Pdf
import fitz  # PyMuPDF

# 全局变量
TEMP_DIR = None
OUTPUT_DIR = None
SUPPORTED_ARCHIVES = ['.zip', '.rar', '.7z', '.tar', '.gz', '.bz2', '.xz']
SUPPORTED_DOCUMENTS = ['.doc', '.docx', '.pdf', '.ppt', '.pptx', '.xls', '.xlsx']


def setup_environment():
    """设置工作环境和目录"""
    global TEMP_DIR, OUTPUT_DIR
    
    # 创建临时目录
    TEMP_DIR = tempfile.mkdtemp(prefix="doc_processor_")
    print(f"创建临时工作目录: {TEMP_DIR}")
    
    # 创建输出目录
    current_dir = os.getcwd()
    OUTPUT_DIR = os.path.join(current_dir, "output")
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    print(f"创建输出目录: {OUTPUT_DIR}")
    
    return TEMP_DIR, OUTPUT_DIR


def cleanup():
    """清理临时文件和目录"""
    if TEMP_DIR and os.path.exists(TEMP_DIR):
        try:
            shutil.rmtree(TEMP_DIR)
            print(f"已清理临时目录: {TEMP_DIR}")
        except Exception as e:
            print(f"清理临时目录时出错: {e}")


def is_archive(filepath):
    """检查文件是否为支持的压缩包格式"""
    return any(filepath.lower().endswith(ext) for ext in SUPPORTED_ARCHIVES)


def is_document(filepath):
    """检查文件是否为支持的文档格式"""
    return any(filepath.lower().endswith(ext) for ext in SUPPORTED_DOCUMENTS)


def extract_archive(archive_path, extract_to):
    """解压文件到指定目录"""
    try:
        print(f"正在解压: {os.path.basename(archive_path)}")
        # 尝试使用pyunpack
        Archive(archive_path).extractall(extract_to)
        return True
    except Exception as e1:
        print(f"pyunpack解压失败，尝试使用patoolib: {e1}")
        try:
            # 尝试使用patoolib
            patoolib.extract_archive(archive_path, outdir=extract_to, verbosity=-1)
            return True
        except Exception as e2:
            print(f"解压文件失败 {archive_path}: {e2}")
            return False


def process_archives_recursively(directory):
    """递归处理目录中的所有压缩文件"""
    found_archives = []
    
    # 查找当前目录中的所有压缩文件
    for root, _, files in os.walk(directory):
        for file in files:
            filepath = os.path.join(root, file)
            if is_archive(filepath):
                found_archives.append(filepath)
    
    # 处理找到的每个压缩文件
    for archive_path in found_archives:
        archive_name = os.path.splitext(os.path.basename(archive_path))[0]
        extract_dir = os.path.join(TEMP_DIR, f"extract_{archive_name}")
        os.makedirs(extract_dir, exist_ok=True)
        
        if extract_archive(archive_path, extract_dir):
            # 递归处理提取出的内容中的压缩文件
            process_archives_recursively(extract_dir)


def convert_word_to_pdf(docx_path, output_path):
    """将Word文档转换为PDF"""
    try:
        from docx2pdf import convert
        print(f"正在转换Word文档: {os.path.basename(docx_path)}")
        convert(docx_path, output_path)
        return os.path.exists(output_path)
    except Exception as e:
        print(f"转换Word文档失败: {e}")
        return False


def convert_ppt_to_pdf(ppt_path, output_path):
    """将PowerPoint转换为PDF"""
    try:
        # 尝试使用COM自动化（Windows）
        import win32com.client
        print(f"正在转换PowerPoint: {os.path.basename(ppt_path)}")
        powerpoint = win32com.client.Dispatch("Powerpoint.Application")
        powerpoint.Visible = True
        
        # 打开并转换
        presentation = powerpoint.Presentations.Open(ppt_path, WithWindow=False)
        presentation.SaveAs(output_path, 32)  # 32 = PDF格式
        presentation.Close()
        powerpoint.Quit()
        
        return os.path.exists(output_path)
    except Exception as e:
        print(f"使用COM自动化转换PowerPoint失败，尝试LibreOffice: {e}")
        try:
            # 尝试使用LibreOffice (跨平台)
            output_dir = os.path.dirname(output_path)
            subprocess.run([
                'soffice',
                '--headless',
                '--convert-to', 'pdf',
                '--outdir', output_dir,
                ppt_path
            ], check=True)
            
            # 检查输出文件
            filename = os.path.splitext(os.path.basename(ppt_path))[0] + '.pdf'
            generated_pdf = os.path.join(output_dir, filename)
            
            # 如果生成的PDF文件名与期望的不同，则重命名
            if generated_pdf != output_path and os.path.exists(generated_pdf):
                os.rename(generated_pdf, output_path)
                
            return os.path.exists(output_path)
        except Exception as e2:
            print(f"使用LibreOffice转换PowerPoint失败: {e2}")
            return False


def convert_excel_to_pdf(xls_path, output_path):
    """将Excel文件转换为PDF"""
    try:
        # 尝试使用COM自动化（Windows）
        import win32com.client
        print(f"正在转换Excel: {os.path.basename(xls_path)}")
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        
        # 打开并转换
        workbook = excel.Workbooks.Open(xls_path)
        workbook.ExportAsFixedFormat(0, output_path)  # 0 = PDF格式
        workbook.Close(False)  # 不保存更改
        excel.Quit()
        
        return os.path.exists(output_path)
    except Exception as e:
        print(f"使用COM自动化转换Excel失败，尝试LibreOffice: {e}")
        try:
            # 尝试使用LibreOffice (跨平台)
            output_dir = os.path.dirname(output_path)
            subprocess.run([
                'soffice',
                '--headless',
                '--convert-to', 'pdf',
                '--outdir', output_dir,
                xls_path
            ], check=True)
            
            # 检查输出文件
            filename = os.path.splitext(os.path.basename(xls_path))[0] + '.pdf'
            generated_pdf = os.path.join(output_dir, filename)
            
            # 如果生成的PDF文件名与期望的不同，则重命名
            if generated_pdf != output_path and os.path.exists(generated_pdf):
                os.rename(generated_pdf, output_path)
                
            return os.path.exists(output_path)
        except Exception as e2:
            print(f"使用LibreOffice转换Excel失败: {e2}")
            return False


def convert_document_to_pdf(doc_path, output_dir):
    """将文档转换为PDF"""
    filename = os.path.basename(doc_path)
    name, ext = os.path.splitext(filename)
    output_path = os.path.join(output_dir, f"{name}.pdf")
    
    if ext.lower() in ['.doc', '.docx']:
        return convert_word_to_pdf(doc_path, output_path)
    elif ext.lower() in ['.ppt', '.pptx']:
        return convert_ppt_to_pdf(doc_path, output_path)
    elif ext.lower() in ['.xls', '.xlsx']:
        return convert_excel_to_pdf(doc_path, output_path)
    elif ext.lower() == '.pdf':
        # 直接复制PDF文件
        try:
            shutil.copy(doc_path, output_path)
            return True
        except Exception as e:
            print(f"复制PDF文件失败: {e}")
            return False
    else:
        print(f"不支持的文档格式: {ext}")
        return False


def find_and_convert_documents(directory, output_dir):
    """查找并转换目录中的所有文档"""
    pdf_files = []
    
    # 遍历目录中的所有文件
    for root, _, files in os.walk(directory):
        for file in files:
            filepath = os.path.join(root, file)
            if is_document(filepath):
                name = os.path.splitext(file)[0]
                output_pdf = os.path.join(output_dir, f"{name}.pdf")
                
                # 转换文档到PDF
                if convert_document_to_pdf(filepath, output_dir):
                    if os.path.exists(output_pdf):
                        pdf_files.append(output_pdf)
    
    return pdf_files


def merge_pdfs(pdf_files, output_path):
    """合并多个PDF文件为一个"""
    if not pdf_files:
        print("没有找到PDF文件，无法合并")
        return False
    
    print(f"正在合并 {len(pdf_files)} 个PDF文件...")
    
    # 使用PyPDF2合并
    try:
        merger = PyPDF2.PdfMerger()
        
        for pdf in pdf_files:
            try:
                merger.append(pdf)
            except Exception as e:
                print(f"添加PDF时出错 ({pdf}): {e}")
        
        # 写入合并的PDF
        merger.write(output_path)
        merger.close()
        
        print(f"成功创建合并的PDF: {output_path}")
        return True
    
    except Exception as e:
        print(f"使用PyPDF2合并PDF失败: {e}")
        
        # 使用pikepdf作为备选方案
        try:
            print("尝试使用pikepdf合并...")
            output_pdf = Pdf.new()
            
            for i, input_file in enumerate(pdf_files):
                try:
                    src = Pdf.open(input_file)
                    output_pdf.pages.extend(src.pages)
                except Exception as e:
                    print(f"处理文件时出错 ({input_file}): {e}")
            
            output_pdf.save(output_path)
            print(f"成功使用pikepdf创建合并的PDF: {output_path}")
            return True
            
        except Exception as e2:
            print(f"使用pikepdf合并失败: {e2}")
            
            # 最后尝试使用PyMuPDF (fitz)
            try:
                print("尝试使用PyMuPDF合并...")
                result = fitz.open()
                
                for pdf in pdf_files:
                    try:
                        with fitz.open(pdf) as doc:
                            result.insert_pdf(doc)
                    except Exception as e:
                        print(f"添加PDF时出错 ({pdf}): {e}")
                
                result.save(output_path)
                print(f"成功使用PyMuPDF创建合并的PDF: {output_path}")
                return True
                
            except Exception as e3:
                print(f"所有合并方法都失败了: {e3}")
                return False


def main():
    """主函数"""
    try:
        # 设置环境
        setup_environment()
        
        # 获取输入目录
        if len(sys.argv) > 1 and os.path.exists(sys.argv[1]):
            input_dir = sys.argv[1]
        else:
            input_dir = os.getcwd()
            
        print(f"处理目录: {input_dir}")
        
        # 处理压缩文件
        process_archives_recursively(input_dir)
        
        # 为PDF文件创建目录
        pdf_dir = os.path.join(TEMP_DIR, "pdfs")
        os.makedirs(pdf_dir, exist_ok=True)
        
        # 查找并转换解压后的文档
        print("\n正在查找和转换文档...")
        converted_pdfs = find_and_convert_documents(TEMP_DIR, pdf_dir)
        
        # 也处理输入目录中的文档
        input_pdfs = find_and_convert_documents(input_dir, pdf_dir)
        
        # 合并所有找到的PDF
        all_pdfs = converted_pdfs + input_pdfs
        
        # 按文件名排序PDF
        all_pdfs.sort()
        
        if all_pdfs:
            print(f"\n找到 {len(all_pdfs)} 个PDF文件")
            output_pdf = os.path.join(input_dir, "合并文档.pdf")
            if merge_pdfs(all_pdfs, output_pdf):
                print(f"\n成功! 合并的PDF已保存到: {output_pdf}")
        else:
            print("\n没有找到可处理的文档")
        
    except Exception as e:
        print(f"处理过程中出错: {e}")
        traceback.print_exc()
    
    finally:
        # 清理临时文件
        cleanup()


if __name__ == "__main__":
    main() 