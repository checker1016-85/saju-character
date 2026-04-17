#!/usr/bin/env python3
"""
엑셀 → JSON 변환 스크립트
━━━━━━━━━━━━━━━━━━━━━━━━
사용법: python excel_to_json.py 사주명리_완전정리_DB.xlsx

엑셀 수정할 때마다 이 스크립트를 실행하면 saju_db.json이 갱신됩니다.
갱신된 JSON을 GitHub에 올리면 웹사이트에 자동 반영됩니다.
"""
import openpyxl, json, sys, os

def convert(xlsx_path):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    
    def read_sheet(name, header_row=2, start_row=3):
        ws = wb[name]
        headers = [str(ws.cell(row=header_row, column=c).value or f'col{c}') for c in range(1, ws.max_column+1)]
        rows = []
        for r in range(start_row, ws.max_row+1):
            row = {}
            for c in range(1, ws.max_column+1):
                v = ws.cell(row=r, column=c).value
                if v: row[headers[c-1]] = str(v)
            if row: rows.append(row)
        return rows

    db = {
        'ilju': read_sheet('시트3_60갑자일주'),
        'wolji': read_sheet('시트4_월주변형레이어'),
        'appearance': read_sheet('시트5_외형옵션카테고리'),
        'gender': read_sheet('시트6_남녀분기표'),
    }
    
    out_path = os.path.join(os.path.dirname(xlsx_path), 'saju_db.json')
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(db, f, ensure_ascii=False, indent=2)
    
    print(f"✅ 변환 완료!")
    print(f"   일주: {len(db['ilju'])}개 / 월지: {len(db['wolji'])}개")
    print(f"   → {out_path}")

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("사용법: python excel_to_json.py 사주명리_완전정리_DB.xlsx")
    else:
        convert(sys.argv[1])
