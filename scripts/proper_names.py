"""為縣市名稱加專名號；只改 run 的單底線，不改正文或其他字元格式。"""
from copy import deepcopy
import re
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

COUNTIES = ['臺北市','台北市','新北市','桃園市','臺中市','台中市','臺南市','台南市','高雄市',
    '基隆市','新竹市','嘉義市','新竹縣','苗栗縣','彰化縣','南投縣','雲林縣','嘉義縣',
    '屏東縣','宜蘭縣','花蓮縣','臺東縣','台東縣','澎湖縣','金門縣','連江縣']
PATTERN = re.compile('|'.join(COUNTIES))


def underline_counties(document):
    changed = 0
    for paragraph in document.element.iter(qn('w:p')):
        nodes = list(paragraph.iter(qn('w:t')))
        text = ''.join(t.text or '' for t in nodes)
        spans = [(m.start(),m.end()) for m in PATTERN.finditer(text)]
        if not spans:
            continue
        offset = 0
        for node in nodes:
            value = node.text or ''
            covered = [any(a<=offset+i<b for a,b in spans) for i in range(len(value))]
            offset += len(value)
            if not any(covered):
                continue
            run = node.getparent()
            if run.tag != qn('w:r') or any(c.tag not in {qn('w:rPr'),qn('w:t')} for c in run) or len(run.findall(qn('w:t')))!=1:
                raise ValueError('縣市名稱包含非純文字 run，請先檢查原稿結構')
            boundaries = [0]+[i for i in range(1,len(value)) if covered[i]!=covered[i-1]]+[len(value)]
            properties = run.find(qn('w:rPr'))
            for start,end in zip(boundaries,boundaries[1:]):
                part = OxmlElement('w:r')
                if properties is not None:
                    part.append(deepcopy(properties))
                if covered[start]:
                    props = part.get_or_add_rPr()
                    underline = props.find(qn('w:u'))
                    if underline is None:
                        underline = OxmlElement('w:u');props.append(underline)
                    underline.set(qn('w:val'),'single')
                t = OxmlElement('w:t');t.text=value[start:end]
                t.set(qn('xml:space'),'preserve');part.append(t)
                run.addprevious(part)
            run.getparent().remove(run)
            changed += 1
    return changed


def format_counties(path):
    document = Document(path)
    if underline_counties(document):
        document.save(path)
