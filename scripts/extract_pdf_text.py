import pdfplumber
import sys
import os

def extract_pdf_text(pdf_path):
    """提取PDF文件的文本内容"""
    text = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for i, page in enumerate(pdf.pages):
                page_text = page.extract_text()
                if page_text:
                    text += f"=== 第{i+1}页 ===\n"
                    text += page_text + "\n\n"
    except Exception as e:
        text = f"读取PDF时出错: {str(e)}\n"
    return text

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: python extract_pdf_text.py <pdf文件路径> [输出文件路径]")
        sys.exit(1)
    
    pdf_path = sys.argv[1]
    if not os.path.exists(pdf_path):
        print(f"文件不存在: {pdf_path}")
        sys.exit(1)
    
    text = extract_pdf_text(pdf_path)
    
    # 如果有输出文件参数，则保存到文件
    if len(sys.argv) >= 3:
        output_path = sys.argv[2]
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(text)
        print(f"文本已保存到: {output_path}")
        print(f"总字符数: {len(text)}")
    else:
        # 否则输出到控制台（截断）
        print(text[:5000])
        print(f"\n... 总字符数: {len(text)}")