import os
import hashlib
import fitz  # PyMuPDF
import pandas as pd
import re
from tqdm import tqdm

def extract_page_text(page):
    """提取页面文本内容"""
    # 使用 get_text("text") 获取纯文本，去除部分格式干扰
    text = page.get_text("text")
    return text.strip()

def get_text_fingerprint(text):
    """
    生成基于前缀的文本指纹（抗噪去重）
    策略：清洗文本 -> 只取前50个字符 -> 生成哈希
    """
    if not text:
        return ""
    
    # 1. 清洗：移除非文字数字字符，只保留核心内容
    # 保留：中文(\u4e00-\u9fff)、英文字母(a-z/A-Z)、数字(0-9)
    cleaned = re.sub(r'[^\u4e00-\u9fff\da-zA-Z]', '', text)
    cleaned = cleaned.lower() # 统一转小写
    
    # 2. 截断：只取前 50 个字符
    # 这是关键！无论后面有多少噪点或页码差异，都只看开头
    prefix = cleaned[:50] 
    
    # 3. 防御性编程：如果清洗后内容太短，返回空（防止全是噪点的页面产生无效指纹）
    if len(prefix) < 5: # 至少要有几个字才算有效
        return ""
        
    return prefix

def is_visually_blank(page, threshold=0.99):
    """
    基于像素分析判断是否为视觉空白页
    threshold: 白色像素占比阈值，默认 99% 以上为白即视为空白
    """
    # 1. 先快速检查是否有文字，有文字肯定不是空白页
    if page.get_text("text").strip():
        return False
        
    # 2. 将页面渲染为低分辨率图片进行分析 (DPI 72 足够判断背景颜色)
    # 矩阵缩放 0.5 倍通常对应 72 DPI (假设原图 144 DPI)
    pix = page.get_pixmap(matrix=fitz.Matrix(0.5, 0.5))
    
    # 获取像素样本
    samples = pix.samples
    total_pixels = len(samples) // 3 # RGB 3个通道
    white_pixels = 0
    
    # 遍历像素 (RGB)
    # samples 是一个字节数组，每3个字节代表一个像素的 R, G, B
    for i in range(0, len(samples), 3):
        r, g, b = samples[i], samples[i+1], samples[i+2]
        # 判断是否为接近白色 (允许轻微的扫描噪点，比如 240,240,240 以上都算白)
        if r > 240 and g > 240 and b > 240:
            white_pixels += 1
            
    white_ratio = white_pixels / total_pixels
    return white_ratio >= threshold

def process_pdf_file(pdf_path):
    """处理单个PDF文件，删除重复和空白页"""
    
    with fitz.open(pdf_path) as doc:
        pages_to_keep = []
        seen_fingerprints = set()
        
        for i in range(doc.page_count):
            page = doc[i]
            raw_text = page.get_text("text")
            
            # --- 1. 检查是否有有效文本 ---
            # 使用更宽松的标准：只要提取出的“核心文字”超过10个，就认为是有内容的
            core_text = re.sub(r'[^\u4e00-\u9fff\da-zA-Z]', '', raw_text)
            has_meaningful_text = len(core_text) > 10
            
            if has_meaningful_text:
                fp = get_text_fingerprint(raw_text)
                
                # 如果能生成指纹，且指纹未见过，则保留
                # 注意：如果指纹为空（比如全是噪点），我们也会保留，防止误删
                if fp and fp in seen_fingerprints:
                    continue # 重复页，跳过
                if fp:
                    seen_fingerprints.add(fp)
                # 如果没有指纹（比如全是噪点），我们保留，但不加入指纹库（防止后续全是噪点的页被误判）
                pages_to_keep.append(i)
                continue
            
            # --- 2. 如果没有有效文本，进行视觉空白检测 ---
            if not is_visually_blank(page, threshold=0.99):
                pages_to_keep.append(i)
        
        # --- 创建新文档 ---
        new_doc = fitz.open()
        try:
            for page_num in pages_to_keep:
                new_doc.insert_pdf(doc, from_page=page_num, to_page=page_num)
            
            temp_path = pdf_path.replace('.pdf', '_cleaned.pdf')
            new_doc.save(temp_path)
            return doc.page_count, len(pages_to_keep), temp_path
        finally:
            new_doc.close()

def process_pdfs_in_directory(directory_path):
    """处理目录下的所有PDF文件"""
    pdf_files = [f for f in os.listdir(directory_path) if f.lower().endswith('.pdf') and '_cleaned' not in f]
    
    results = []
    
    # 这里的 tqdm 用于显示进度条
    for idx, filename in enumerate(tqdm(pdf_files, desc="处理中")):
        file_path = os.path.join(directory_path, filename)
        
        try:
            original_pages, processed_pages, cleaned_path = process_pdf_file(file_path)
            results.append({
                '序号': idx + 1,
                '文件名': filename,
                '处理前页数': original_pages,
                '处理后页数': processed_pages
            })
        except Exception as e:
            print(f"\n处理文件 {filename} 时出错: {str(e)}")
            # 出错也记录，防止中断
            try:
                doc = fitz.open(file_path)
                results.append({'序号': idx + 1, '文件名': filename, '处理前页数': doc.page_count, '处理后页数': '错误'})
                doc.close()
            except:
                results.append({'序号': idx + 1, '文件名': filename, '处理前页数': '?', '处理后页数': '错误'})

    output_excel_path = os.path.join(directory_path, "处理列表.xlsx")
    df = pd.DataFrame(results)
    df.to_excel(output_excel_path, index=False)
    print(f"\n完成！结果保存在: {output_excel_path}")

def main():
    print("PDF 智能清理工具 (去空白/去重)")
    directory_path = input("请输入PDF目录路径: ").strip().strip('"')
    if not os.path.isdir(directory_path):
        print("目录不存在！")
        return
    process_pdfs_in_directory(directory_path)

if __name__ == "__main__":
    main()