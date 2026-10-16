"""檢查領域專家欄位、批註與追蹤修訂；不以數據／圖片 QA 代替內容完成度。"""
import argparse
import json
from pathlib import Path
import re
from zipfile import ZipFile

from lxml import etree
from inspect_reference_docs import NS, W, text, question_blocks


def pending(value):
    if re.search(r'待(?:補|填|確認)|本題(?:主要|為|旨在)\s*[.．…]+|顯示\s*[.．…]+',value):return True
    stripped=re.sub(r'[（(][一二三四五六七八九十\d]+[）)]','',value)
    return not re.search(r'[A-Za-z\u3400-\u9fff]',stripped)


def section(content,start,end=None):
    match=re.search(start,content)
    if not match:return ''
    value=content[match.end():]
    if end:
        following=re.search(end,value)
        if following:value=value[:following.start()]
    return value.strip()


def verify(path):
    path=Path(path)
    with ZipFile(path) as archive:
        doc=etree.fromstring(archive.read('word/document.xml'))
        comments=0
        if 'word/comments.xml' in archive.namelist():
            comments=len(etree.fromstring(archive.read('word/comments.xml')).findall(W+'comment'))
        revisions=len(doc.xpath('.//w:ins|.//w:del',namespaces=NS))
        items=[]
        for table_index,table in enumerate(doc.xpath('./w:body/w:tbl',namespaces=NS)):
            for row_start,rows,metadata in question_blocks(table):
                # 只讀取目前這一題，不讓 Word 合併的下一題補滿本題空欄。
                content='\n'.join(text(p) for row in rows for p in row.xpath('.//w:p',namespaces=NS))
                compact=re.sub(r'\s+','',content)
                labels={'assessment_target':(r'[（(]一[）)]評量目標',r'[（(]二[）)]數據說明'),
                    'option_analysis':(r'[（(]三[）)]選項及學生表現說明',r'二、對應學習重點'),
                    'learning_alignment':(r'二、對應學習重點',r'三、教學建議'),
                    'teaching_suggestions':(r'三、教學建議',None)}
                fields={name:section(compact,*bounds) for name,bounds in labels.items()}
                focus=rows[2].findall(W+'tc') if len(rows)>2 else []
                fields['learning_focus']=''.join(text(c) for c in focus[1:])
                missing=[name for name,value in fields.items() if pending(value)]
                items.append({**metadata,'table_index':table_index,'row_start':row_start,
                    'pending_fields':missing,'complete':not missing})
        duplicates=sorted({i['code'] for i in items if sum(j['code']==i['code'] for j in items)>1})
        return {'file':path.name,'questions':len(items),'complete_questions':sum(i['complete'] for i in items),
            'pending_questions':sum(not i['complete'] for i in items),'items':items,
            'comments':comments,'tracked_changes':revisions,'duplicate_codes':duplicates,
            'editorial_ready':bool(items) and all(i['complete'] for i in items) and not comments and not revisions and not duplicates,
            'note':'完整度通過僅代表欄位有內容；題意、學習指標、選項迷思與教學建議仍須領域審核。'}


if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--dir',type=Path,required=True)
    parser.add_argument('--json',type=Path,required=True)
    parser.add_argument('--require-complete',action='store_true',help='將待補欄位、批註、修訂視為交付阻擋')
    args=parser.parse_args()
    reports=[verify(p) for p in sorted(args.dir.glob('第四章*.docx'))]
    args.json.write_text(json.dumps(reports,ensure_ascii=False,indent=2),encoding='utf-8')
    print([(r['file'],r['questions'],r['pending_questions'],r['editorial_ready']) for r in reports])
    if not reports or (args.require_complete and any(not r['editorial_ready'] for r in reports)):raise SystemExit(1)
