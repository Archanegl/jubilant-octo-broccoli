"""generate_pptx.py

将仓库根目录中名为 presentation.pptx 的文本草稿（Markdown 风格）解析为真实的 PowerPoint .pptx 文件。

用法（在本地）：
1. pip install python-pptx
2. python generate_pptx.py

生成的文件： output/presentation_real.pptx

脚本说明：
- 解析包含多个“幻灯片 N — 标题”块的文本文件。每个块中的 "-" 列表项被视为幻灯片要点。
- 以“备注：”开头的行会被当作该幻灯片的备注（插入到幻灯片备注区）。
- 在每张幻灯片的内容区底部加入一行占位文字“[示意图占位 - 在此插入图片，建议来源：示意图/光纤/雷达 图片]”。

注意：此脚本在 GitHub Actions 中运行也可用（已附带示例 workflow），workflow 会将生成的文件提交回仓库。
"""

from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT
import re
import os

INPUT_PATH = 'presentation.pptx'  # 当前仓库中保存的草稿文本文件
OUTPUT_DIR = 'output'
OUTPUT_FILE = os.path.join(OUTPUT_DIR, 'presentation_real.pptx')

SLIDE_TITLE_RE = re.compile(r'^幻灯片\s*\d+\s*[—-]\s*(.+)$')
NOTE_PREFIX = '备注：'


def parse_draft(path):
    """解析草稿文本，返回 slides 列表：[{title: str, bullets: [str], note: str}]"""
    with open(path, 'r', encoding='utf-8') as f:
        lines = [ln.rstrip('\n') for ln in f]

    slides = []
    cur = None

    for ln in lines:
        m = SLIDE_TITLE_RE.match(ln)
        if m:
            # start new slide
            if cur:
                slides.append(cur)
            cur = {'title': m.group(1).strip(), 'bullets': [], 'note': ''}
            continue
        if cur is None:
            continue
        # bullet lines start with - 或者以中文短横开头
        stripped = ln.strip()
        if stripped.startswith('- '):
            cur['bullets'].append(stripped[2:].strip())
            continue
        if stripped.startswith('备注：') or stripped.startswith(NOTE_PREFIX):
            # collect remainder as note (append if multi-line note)
            note_text = stripped[len(NOTE_PREFIX):].strip()
            if cur['note']:
                cur['note'] += '\n' + note_text
            else:
                cur['note'] = note_text
            continue
        # 有些行可能是纯要点文本（没有前导 -），尝试以中文序号或其他形式识别
        if stripped and not stripped.startswith('---'):
            # treat as continuation of previous bullet if exists, else as new bullet
            if cur['bullets']:
                cur['bullets'][-1] += ' ' + stripped
            else:
                cur['bullets'].append(stripped)

    if cur:
        slides.append(cur)
    return slides


def make_pptx(slides, out_path):
    prs = Presentation()
    # use a clean slide layout: title and content (usually layout 1)
    title_layout = prs.slide_layouts[0]
    content_layout = prs.slide_layouts[1]

    # cover slide from first slide content
    if slides:
        s0 = prs.slides.add_slide(title_layout)
        s0.shapes.title.text = slides[0]['title']
        if slides[0]['bullets']:
            # put first bullet as subtitle
            subtitle = s0.placeholders[1]
            subtitle.text = '\n'.join(slides[0]['bullets'][:3])
        # notes
        notes = s0.notes_slide.notes_text_frame
        if slides[0]['note']:
            notes.text = slides[0]['note']

    # subsequent slides
    for sl in slides[1:]:
        slide = prs.slides.add_slide(content_layout)
        slide.shapes.title.text = sl['title']
        # add bullets
        body = slide.shapes.placeholders[1].text_frame
        body.clear()
        for i, b in enumerate(sl['bullets']):
            p = body.add_paragraph() if i>0 else body.paragraphs[0]
            p.text = b
            p.level = 0
            p.font.size = Pt(18)
        # add image placeholder text box at bottom
        left = Inches(0.5)
        top = Inches(5.0)
        width = Inches(9)
        height = Inches(0.6)
        txBox = slide.shapes.add_textbox(left, top, width, height)
        tf = txBox.text_frame
        tf.text = '[示意图占位 - 在此插入图片，建议来源：示意图/光纤/雷达 图片]'
        for para in tf.paragraphs:
            para.font.italic = True
            para.font.size = Pt(12)
            para.alignment = PP_PARAGRAPH_ALIGNMENT.CENTER
        # notes
        notes = slide.notes_slide.notes_text_frame
        if sl['note']:
            notes.text = sl['note']

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    prs.save(out_path)
    print('Saved PPTX to', out_path)


if __name__ == '__main__':
    if not os.path.exists(INPUT_PATH):
        print(f'Error: input file not found: {INPUT_PATH}')
    else:
        slides = parse_draft(INPUT_PATH)
        if not slides:
            print('No slides parsed from draft. Please check the input file format.')
        else:
            make_pptx(slides, OUTPUT_FILE)
