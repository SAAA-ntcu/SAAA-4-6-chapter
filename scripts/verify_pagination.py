"""核對 Word 實際轉出的 PDF：每題選答率與試題分析開頭必須同頁。"""
import argparse
import json
from pathlib import Path
import re

from docx import Document
import pymupdf
from report_images import item_tables


def verify(docx_path,pdf_path):
    expected={table.cell(1,1).text.strip() for table in item_tables(Document(docx_path)).values()}
    found=[]
    errors=[]
    articles=[]
    audit_path=Path(docx_path).with_suffix('.images.json')
    audit=json.loads(audit_path.read_text(encoding='utf-8')) if audit_path.exists() else {}
    with pymupdf.open(pdf_path) as document:
        for number,page in enumerate(document,1):
            text=re.sub(r'\s+','',page.get_text())
            codes=re.findall(r'\d{3}-[CMES]\d-[A-Z]\d-\d{2}',text)
            if len(codes)>1:errors.append(f'第{number}頁有多題起始表格')
            if codes:found.append({'code':codes[0],'question_page':number,'statistics_pages':[],'analysis_pages':[]})
            if found:
                if '選答率' in text:found[-1]['statistics_pages'].append(number)
                if '一、試題分析' in text:found[-1]['analysis_pages'].append(number)
            if len(text)<5 and not page.get_images():errors.append(f'第{number}頁為空白頁')
        pages=len(document)
        if audit.get('chinese_layout')=='intro_with_first_item':
            for label in ['一、語文知識：','二、閱讀理解：']:
                matches=[p for p in document if label in re.sub(r'\s+','',p.get_text())]
                if len(matches)!=1 or not re.search(r'\d{3}-C\d-[A-Z]\d-\d{2}',matches[0].get_text()):
                    errors.append(label+' 說明與首題表格未同頁')
        for material in audit.get('material_layouts',[]):
            if material.get('mode')!='single_page_composite':continue
            first,last=material['question_range']
            label=f'第{first}～{last}題共用文章／材料'
            matches=[(i,p) for i,p in enumerate(document,1) if label in re.sub(r'\s+','',p.get_text())]
            valid=False
            if len(matches)==1:
                number,page=matches[0]
                images=page.get_image_info()
                if len(images)==1:
                    box=pymupdf.Rect(images[0]['bbox'])
                    width,height=box.width/72*2.54,box.height/72*2.54
                    valid=(abs(width-material['width_cm'])<.04 and abs(height-material['height_cm'])<.04
                           and page.rect.contains(box))
                    if material.get('vertical_alignment')=='center' and material.get('target_center_y_cm') is not None:
                        actual_center=(box.y0+box.y1)/2/72*2.54
                        centered=abs(actual_center-material['target_center_y_cm'])<.2
                        if not centered:errors.append(material['id']+' 未在文章區域垂直置中')
                        valid=valid and centered
                articles.append({'id':material['id'],'page':number,'source_parts':material['part_count'],
                                 'single_page':valid})
            if not valid:errors.append(material['id']+' 整篇文章未完整放在一頁內')
    if {r['code'] for r in found}!=expected or len(found)!=len(expected):
        errors.append('PDF 試題代碼／題數與 Word 不一致')
    for record in found:
        record['connected']=bool(set(record['statistics_pages'])&set(record['analysis_pages']))
        if not record['connected']:errors.append(record['code']+' 的選答率與分析開頭分頁')
    return {'file':Path(docx_path).name,'pages':pages,'questions':found,'articles':articles,'passed':not errors,'errors':errors}


if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--dir',required=True,type=Path)
    parser.add_argument('--pdf-dir',required=True,type=Path)
    parser.add_argument('--json',required=True,type=Path)
    args=parser.parse_args()
    results=[verify(path,args.pdf_dir/(path.stem+'.pdf')) for path in sorted(args.dir.glob('第四章*.docx'))]
    args.json.write_text(json.dumps(results,ensure_ascii=False,indent=2),encoding='utf-8')
    print([(r['file'],len(r['questions']),r['passed'],r['errors']) for r in results])
    if not results or any(not r['passed'] for r in results):raise SystemExit(1)
