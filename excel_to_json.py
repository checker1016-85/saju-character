"""
사주명리_완전정리_DB.xlsx → saju_db.json 변환기
GitHub Actions 자동화용. DB 수정 후 재변환.
pip install pandas openpyxl
"""
import pandas as pd, json, sys, os

XLSX = os.path.join(os.path.dirname(os.path.abspath(__file__)), '사주명리_완전정리_DB.xlsx')
if len(sys.argv) > 1:
    XLSX = sys.argv[1]

COL12 = {
    4:'A1_골격',5:'A2_체형',6:'B1_얼굴형',7:'B2_눈',8:'B3_이면눈매',
    9:'B4_눈썹',10:'B5_코',11:'B6_입',12:'B7_얼굴음영',
    13:'C1_피부톤',14:'C2_오행변이',
    15:'D1_헤어',16:'D2_머리물리',17:'D3_수염',
    18:'E1_의복핏',19:'E2_상의',20:'E3_하의',21:'E4_풀바디',22:'E5_신발',
    23:'F1_머리악세',24:'F2_몸악세',25:'F3_소울웨폰',26:'F4_장신구공명',
    27:'F5_인장문양',28:'F6_특수요소',
    29:'G1_아우라',30:'G2_이펙트컬러',31:'G3_그림자',32:'G4_발걸음',
    33:'G5_동작잔상',34:'G6_지장간',
    35:'H1_표정온도',36:'H2_분위기',37:'H3_제스처',38:'H4_음성톤',
    39:'H5_감정중력',40:'H6_각성트리거',41:'H7_나이대종족'
}

def cell(df,r,c):
    v=df.iloc[r,c]
    return str(v).strip() if pd.notna(v) and str(v).strip() else ''

def extract_dna(xlsx):
    df=pd.read_excel(xlsx,sheet_name='시트12_일주_캐릭터DNA',header=None)
    dna={}; cur=None
    for i in range(3,len(df)):
        ilju=cell(df,i,1)
        if ilju: cur=ilju
        if not cur: continue
        g={'共':'common','♂':'male','♀':'female'}.get(cell(df,i,3))
        if not g: continue
        if cur not in dna: dna[cur]={}
        rd={}
        for ci,cn in COL12.items():
            v=cell(df,i,ci)
            if v: rd[cn]=v
        dna[cur][g]=rd
    return dna

def extract_gender(xlsx):
    df=pd.read_excel(xlsx,sheet_name='시트10_성별',header=None)
    gt={}
    cols={3:'성격분기',4:'외형분기',5:'체형분기',6:'피부인상분기',
          7:'눈빛표정분기',8:'분위기태그',9:'스타일분기',
          10:'직업방향',11:'관계패턴',12:'프롬프트외형태그'}
    for i in range(2,len(df)):
        cg=cell(df,i,1); gd=cell(df,i,2)
        if not cg or not gd: continue
        hj=cg.split('(')[1].split(')')[0] if '(' in cg else cg
        key=f"{hj}_{gd}"
        rd={}
        for ci,cn in cols.items():
            v=cell(df,i,ci)
            if v: rd[cn]=v
        gt[key]=rd
    return gt

def extract_ilju(xlsx):
    df=pd.read_excel(xlsx,sheet_name='시트11_일주',header=None)
    iu={}
    cols={3:'천간',4:'지지',5:'납음',6:'납음해석',7:'십이운성',
          8:'일지십성',9:'생극관계',10:'고유성격',11:'고유외형',12:'고유이미지'}
    for i in range(2,len(df)):
        ilju=cell(df,i,2)
        if not ilju: continue
        rd={}
        for ci,cn in cols.items():
            v=cell(df,i,ci)
            if v: rd[cn]=v
        iu[ilju]=rd
    return iu

def extract_sipsung_jobs(xlsx):
    df=pd.read_excel(xlsx,sheet_name='시트6_지지십성',header=None)
    sj={}
    for i in range(2,len(df)):
        name=cell(df,i,1)
        job=cell(df,i,6)
        if name and job:
            sj[name]=job
    return sj

def extract_job_visual(xlsx):
    try:
        df=pd.read_excel(xlsx,sheet_name='시트14_직업비주얼',header=None)
    except:
        return {}
    jv={}
    for i in range(2,len(df)):
        name=cell(df,i,1)
        desc=cell(df,i,2)
        sd=cell(df,i,3)
        if name:
            entry={}
            if desc: entry['desc']=desc
            if sd: entry['sd']=sd
            jv[name]=entry
    return jv

if __name__=='__main__':
    print(f'Reading {XLSX}...')
    result={
        "dna": extract_dna(XLSX),
        "gender_traits": extract_gender(XLSX),
        "ilju_unique": extract_ilju(XLSX),
        "sipsung_jobs": extract_sipsung_jobs(XLSX),
        "job_visual": extract_job_visual(XLSX)
    }
    out=os.path.join(os.path.dirname(XLSX) or '.','saju_db.json')
    with open(out,'w',encoding='utf-8') as f:
        json.dump(result,f,ensure_ascii=False,separators=(',',':'))
    sz=os.path.getsize(out)/1024
    print(f'Done -> {out} ({sz:.0f}KB)')
    print(f'  DNA: {len(result["dna"])}')
    print(f'  Gender: {len(result["gender_traits"])}')
    print(f'  Ilju: {len(result["ilju_unique"])}')
    print(f'  SipsungJobs: {len(result["sipsung_jobs"])}')
    print(f'  JobVisual: {len(result["job_visual"])}')
