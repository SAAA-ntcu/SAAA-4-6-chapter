"""題本 PDF＋橫式答案＋評量架構 PPTX＋施測統計，整合產出第四、六章。"""
import argparse
import json
import hashlib
import html
import re
from pathlib import Path

from generate_reports import generate_chapter_4, generate_chapter_6, read_analysis_file, resolve_templates
from question_screenshots import export, identity
from source_inputs import discover_source_sets, missing_statistics
from verify_editorial import verify as verify_editorial


def build(input_dir, analysis_dir, output_dir, year='115', grade='7N', suffix='_AI產出', draft=False, crop_overrides=None, subjects=None):
    input_dir, output_dir = Path(input_dir).resolve(), Path(output_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    template_dir = Path(__file__).resolve().parents[1]/'resources/templates'
    results = []
    overrides = json.loads(Path(crop_overrides).read_text(encoding='utf-8')) if crop_overrides else {}
    for data in discover_source_sets(input_dir):
        subject, answers, domains = data['subject'], data['answers'], data['domains']
        if subjects and subject not in subjects:
            continue
        cohort = subject+grade
        analysis_file = Path(analysis_dir)/f'{cohort}_試題分析總表.xlsx' if analysis_dir else None
        if analysis_file and analysis_file.exists():
            datasets = read_analysis_file(str(analysis_file))
            # 單科新生檔的工作表可能仍名為「7年級」，不任取其他年級。
            analysis = datasets.get(grade) or datasets.get(re.sub(r'\D','',grade))
            if analysis is None:
                candidates = {id(v):v for v in datasets.values()}
                if len(candidates) != 1:
                    raise ValueError(f'{analysis_file} 無法唯一辨認年級 {grade}')
                analysis = next(iter(candidates.values()))
        elif draft:
            analysis = missing_statistics(answers)
        else:
            raise FileNotFoundError(f'缺少 {cohort}_試題分析總表.xlsx；請提供 --analysis-dir，或明確使用 --draft')
        numbers = [q['num'] for q in analysis['questions']]
        if sorted(numbers) != list(range(1,len(answers)+1)):
            raise ValueError(f'{cohort} 統計／答案題號不一致或重複')
        for q in analysis['questions']:
            if not q['invalid'] and q['correct_ans'] != answers[q['num']-1]:
                raise ValueError(f'{cohort} q{q["num"]} 統計答案 {q["correct_ans"]} 與正式答案 {answers[q["num"]-1]} 不一致')
        crops = output_dir/'題本截圖_AI產出'
        crop_result = export(data['booklet'], 'booklets', crops, 220, overrides, {})
        manifest = crops/'booklets'/identity(data['booklet'])/'manifest.json'
        if crop_result['images'] != len(answers):
            raise ValueError(f'{cohort} 截圖題數與正式答案不一致')
        t4, t6 = resolve_templates(str(template_dir), subject)
        c4, c6 = output_dir/f'第四章_{cohort}{suffix}.docx', output_dir/f'第六章_{cohort}{suffix}.docx'
        generate_chapter_4(subject, grade, year, answers, domains, analysis, t4, str(c4), manifest)
        generate_chapter_6(subject, grade, year, answers, domains, analysis, t6, str(c6))
        editorial=verify_editorial(c4)
        (output_dir/f'專家欄位檢驗_{cohort}{suffix}.json').write_text(
            json.dumps(editorial,ensure_ascii=False,indent=2),encoding='utf-8')
        sources = [data['booklet'], data['answer_file'], data['dimension_file']]
        if analysis_file and analysis_file.exists():
            sources.append(analysis_file)
            raw = analysis_file.with_name(f'{cohort}_分析結果.xlsx')
            if raw.exists():
                sources.append(raw)
        result = {'cohort':cohort, 'draft':bool(analysis.get('draft')), 'questions':len(answers),
            'editorial_ready':editorial['editorial_ready'], 'pending_expert_questions':editorial['pending_questions'],
            'chapter4':str(c4), 'chapter6':str(c6), 'manifest':str(manifest),
            'answers':answers, 'domains':domains, 'source_cells':data['provenance'],
            'analysis_file':str(analysis_file) if analysis_file else None,
            'sources':[{'path':str(p.resolve()), 'sha256':hashlib.sha256(p.read_bytes()).hexdigest()} for p in sources]}
        results.append(result)
        (output_dir/f'輸入對照_{cohort}{suffix}.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    (output_dir/f'產出清單{suffix}.json').write_text(json.dumps(results,ensure_ascii=False,indent=2),encoding='utf-8')
    links = []
    for result in results:
        preview = Path(result['manifest']).parent/'index.html'
        status='專家欄位已填，仍須領域審核' if result['editorial_ready'] else f'專家欄位待補：{result["pending_expert_questions"]} 題'
        links.append('<tr><td>'+html.escape(result['cohort'])+'</td><td>'+str(result['questions'])+'</td><td>'+status+'</td>'+''.join(
            '<td><a href="'+html.escape(str(p.relative_to(output_dir)).replace('\\','/'))+'">'+label+'</a></td>'
            for p,label in [(Path(result['chapter4']),'第四章'),(Path(result['chapter6']),'第六章'),(preview,'題本圖片與文章')])+'</tr>')
    (output_dir/f'報告索引{suffix}.html').write_text('<!doctype html><meta charset="utf-8"><title>報告產出索引</title><style>body{font:18px system-ui;margin:40px}td,th{padding:15px;border-bottom:1px solid #ddd}a{color:#145a92}</style><h1>報告與題本圖片</h1><table><tr><th>科目</th><th>題數</th><th>內容完成度</th><th>第四章</th><th>第六章</th><th>截圖</th></tr>'+''.join(links)+'</table>',encoding='utf-8')
    return results


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--input-dir', required=True, type=Path)
    parser.add_argument('--analysis-dir', type=Path)
    parser.add_argument('--out-dir', required=True, type=Path)
    parser.add_argument('--grade', default='7N')
    parser.add_argument('--year', default='115')
    parser.add_argument('--output-suffix', default='_AI產出')
    parser.add_argument('--draft', action='store_true')
    parser.add_argument('--crop-overrides',type=Path,help='自訂裁切座標 JSON，以來源完整檔名為鍵')
    parser.add_argument('--subject',choices=['C','M','E'],help='僅重建指定科目，未指定時產出全部科目')
    args = parser.parse_args()
    build(args.input_dir,args.analysis_dir,args.out_dir,args.year,args.grade,args.output_suffix,args.draft,args.crop_overrides,
          subjects=[args.subject] if args.subject else None)


if __name__=='__main__':
    main()
