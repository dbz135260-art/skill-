import docx
import sys
import os

def extract_doc_text(doc_path):
    """提取DOC文件的文本内容"""
    text = ""
    try:
        doc = docx.Document(doc_path)
        for i, para in enumerate(doc.paragraphs):
            if para.text.strip():
                text += para.text + "\n"
    except Exception as e:
        text = f"读取DOC时出错: {str(e)}\n"
    return text

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: python extract_doc_text.py <doc文件路径> [输出文件路径]")
        sys.exit(1)
    
    doc_path = sys.argv[1]
    if not os.path.exists(doc_path):
        print(f"文件不存在: {doc_path}")
        sys.exit(1)
    
    text = extract_doc_text(doc_path)
    
    if len(sys.argv) >= 3:
        output_path = sys.argv[2]
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(text)
        print(f"文本已保存到: {output_path}")
        print(f"总字符数: {len(text)}")
    else:
        print(text[:5000])
        print(f"\n... 总字符数: {len(text)}")
