import re
from pypinyin import pinyin, Style
from docx import Document

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>带拼音儿童读物（拼音在上）</title>
    <style>

        /* 核心样式 */
        .pinyin-container {{
            max-width: 800px;
            margin: 20px auto;
            padding: 30px;
            line-height: 1.5em;  /* 减少行高使更紧凑 */
            font-family: "华文楷体", SimSun;
            font-size: 20px;
            background: #f9f9f9;
            border-radius: 10px;
        }}

        /* 标题样式 */
        .title {{
            font-size: 28px;
            text-align: center;
            margin: 1.5em 0;
            line-height: 1.8 !important;
        }}
        .title ruby rt {{
            font-size: 0.5em;  /* 标题拼音相对大小 */
            padding-bottom: 0.3em;
        }}
        .title ruby rb {{
            padding-top: 0.2em;
        }}

        /* 段落首行缩进 */
        .pinyin-container p {{
            text-indent: 2em;  /* 中文标准缩进 */
            margin: 0.8em 0;
        }}
        ruby {{
            ruby-position: over;
            -webkit-ruby-position: before;
            display: inline-flex;
            flex-direction: column-reverse;
            vertical-align: middle;  /* 关键调整点 */
            margin-right: 0.3em;
        }}

        rt {{
            font-size: 0.7em;
            color: #e74c3c;
            letter-spacing: normal;
            padding-bottom: 0em;  /* 减少拼音底部间距 */
            order: 1;
        }}

        rb {{
            order: 2;
            padding-top: 0em;  /* 减少文字顶部间距 */
            vertical-align: middle;  /* 文字垂直居中 */
        }}

        /* 标点符号精准对齐 */
        .punctuation {{
            
            vertical-align: bottom;  /* 与文字对齐 */
            margin-right: 0.01em;
            position: relative;
            top: 0.1em;  /* 微调位置 */
        }}



        /* 响应式适配 */
        @media (max-width: 600px) {{
            .pinyin-container {{
                font-size: 16px;
                line-height: 1.5em;
                padding: 15px;
            }}
            .title {{
                font-size: 22px;
                line-height: 1.6 !important;
            }}
            .pinyin-container p {{
                text-indent: 2em;  /* 保持缩进 */
            }}
            rt {{
                font-size: 0.5em;
                padding-bottom: 0em;
            }}
            .punctuation {{
                top: 0.05em;
                margin-right: 0.01em;
            }}
        }}

        /* 打印优化 */
        @media print {{
            .pinyin-container {{
                background: none;
                padding: 0;
            }}
            rt {{ color: black !important; }}
        }}
    </style>
</head>
<body>
    <div class="pinyin-container">
        {content}
    </div>
        <script>
        // 动态调整拼音间距
        document.querySelectorAll('rt').forEach(rt => {{
            const pyLength = rt.textContent.length;
            const spaceSpan = rt.closest('ruby').nextElementSibling;
            if(spaceSpan) {{
                spaceSpan.style.setProperty('--py-len', pyLength);
            }}
        }});
    </script>
</body>
</html>
"""
# 指定要处理的Word文档路径
doc_path = 'C:/Users/lilil/Desktop/yqzddwj.docx'
output_path = 'C:/Users/lilil/Desktop/annotated_yqzddwj.html'

def generate_pinyin_html(doc_path, output_file=output_path):
    # 中文标点正则表达式（扩展版）
    chinese_punct = re.compile(r'[\u3000-\u303F\uFF00-\uFFEF\u201C-\u201D\u2018-\u2019]')
    # 打开现有的Word文档
    try:
        doc = Document(doc_path)
    except Exception as e:
        print(f"无法打开文档: {e}")
        exit(1)

    # def is_chinese_punctuation(char):
    #     """识别中文标点正则表达式"""
    #     return re.match(r'[\u3000-\u303F\uff00-\uffef]', char)

    html_content = []
    # 遍历文档中的所有段落
    for paragraph in doc.paragraphs:          
       
        para = paragraph.text.strip()
        if not para:
            continue
        pys = pinyin(para, style=Style.TONE, heteronym=False)
        line_html = []
        
        for char, py in zip(para, pys):
            if '\u4e00' <= char <= '\u9fff':  # 判断是否为汉字
                # 获取单个汉字的拼音列表（可能有多个读音）
                pinyin_list = pinyin(char, style=Style.TONE)
                # 取第一个拼音（对于多音字，这里只取最常见的一种）
                pinyin_str = pinyin_list[0][0] if pinyin_list else ''
                line_html.append(f'<ruby><rt>{pinyin_str}</rt>{char}</ruby>')
            else:
                line_html.append(f'<span class="punctuation">{char}</span>')

        # 若没有应用了“正文”样式，则添加title样式
        if paragraph.style.name != 'Normal':      
            html_content.append(f'<h1 class="title">{"".join(line_html)}</h1>')
        else:
            html_content.append(f'<p>{"".join(line_html)}</p>')

    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(HTML_TEMPLATE.format(content='\n'.join(html_content)))

if __name__ == "__main__":    
    generate_pinyin_html(doc_path)
    print("HTML文件生成成功！")

""" # 保存新的文档到指定路径
try:
    new_doc.save(output_path)
    print(f"文档已成功保存至: {output_path}")
except Exception as e:
    print(f"无法保存文档: {e}")
 """