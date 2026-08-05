# -*- coding: utf-8 -*-
import json
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SC="/tmp/claude-0/-home-user-jooho96/2e90a833-c170-58f0-a6f1-86983b335b93/scratchpad"
P=json.load(open(f"{SC}/poland.json",encoding="utf-8"))
C=json.load(open(f"{SC}/customers2.json",encoding="utf-8"))
FONT="맑은 고딕"
XLSX="/home/user/jooho96/현대로템_분기별_단일판매공급계약_매출인식.xlsx"
wb=openpyxl.load_workbook(XLSX)

navy=PatternFill("solid",fgColor="1F3864"); blue=PatternFill("solid",fgColor="2E5496")
defe=PatternFill("solid",fgColor="FCE4D6"); ann=PatternFill("solid",fgColor="FFF2CC"); grand=PatternFill("solid",fgColor="BDD7EE")
hf=Font(name=FONT,bold=True,color="FFFFFF",size=10); bf=Font(name=FONT,bold=True,size=10)
nf=Font(name=FONT,size=10); sf=Font(name=FONT,size=8,color="808080")
thin=Side(style="thin",color="BFBFBF"); bd=Border(left=thin,right=thin,top=thin,bottom=thin)
ctr=Alignment("center","center",wrap_text=True); lft=Alignment("left","center",wrap_text=True); rgt=Alignment("right","center")
WON="#,##0;(#,##0);-"; PCT="+0.0%;-0.0%;-"

# ============ 폴란드 추적 시트 ============
ws=wb.create_sheet("폴란드_방산고객_분기추적",1)
ncol=7
ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=ncol)
c=ws.cell(1,1,"현대로템 연결 최대고객 = 폴란드(방산)  분기별 매출 추적"); c.font=Font(name=FONT,bold=True,color="FFFFFF",size=14); c.fill=navy; c.alignment=lft; ws.row_dimensions[1].height=26
ws.merge_cells(start_row=2,start_column=1,end_row=2,end_column=ncol)
c=ws.cell(2,1,"출처: 연결재무제표 주석 '매출 10% 이상 외부 고객'(2024~ '주요 고객에 대한 공시')  ·  단위: 백만원  ·  폴란드=K2전차(1·2차) 발주처. 단일판매 K2 당기와 정확히 일치 검증됨.")
c.font=sf; c.alignment=lft; ws.row_dimensions[2].height=15
cols=["보고서","구분","폴란드 방산매출\n(당분기 3개월)","폴란드 누적\n(YTD/연간)","전분기比\n(3개월)","고객 표기","비고"]
HDR=3
for j,h in enumerate(cols,1):
    cc=ws.cell(HDR,j,h); cc.font=hf; cc.fill=navy; cc.alignment=ctr; cc.border=bd
ws.row_dimensions[HDR].height=32
PTYPE={"2023Q2":"반기","2023Q3":"분기","2023Q4":"사업(연간)","2024Q1":"분기","2024Q2":"반기","2024Q3":"분기",
 "2024Q4":"사업(연간)","2025Q1":"분기","2025Q2":"반기","2025Q3":"분기","2025Q4":"사업(연간)","2026Q1":"분기"}
IDXLBL={"2023Q2":"고객2(방산·신규)","2023Q3":"고객2(방산)","2023Q4":"고객2","2024Q1":"고객1","2024Q2":"고객1","2024Q3":"고객1",
 "2024Q4":"고객1","2025Q1":"고객1","2025Q2":"고객1","2025Q3":"고객1","2025Q4":"고객1","2026Q1":"고객1"}
NOTE={"2023Q2":"연결매출 10%↑ 첫 등장(K2 폴란드 인도 개시)","2024Q1":"폴란드가 최대고객(고객1)으로 등극",
 "2023Q4":"연간 3개월=연간−9M 계산","2024Q4":"연간 3개월=연간−9M 계산","2025Q4":"연간 3개월=연간−9M 계산"}
ORD=P["ord"]; m3=P["m3"]; ytd=P["ytd"]
r=HDR+1; prev3=None
for lab in ORD:
    v3=m3[lab]/1000.0; vy=ytd[lab]/1000.0
    annual = PTYPE[lab].startswith("사업")
    fill=ann if annual else None
    def put(j,v,fmt=None,al=None,font=nf):
        cc=ws.cell(r,j,v); cc.font=font; cc.border=bd; cc.alignment=al or (rgt if isinstance(v,(int,float)) else ctr)
        if fmt: cc.number_format=fmt
        if fill: cc.fill=fill
    put(1,lab); put(2,PTYPE[lab])
    put(3,round(v3),WON,font=bf); put(4,round(vy),WON)
    qoq=(v3/prev3-1) if (prev3 and prev3!=0) else None
    put(5,qoq,PCT)
    put(6,IDXLBL[lab])
    put(7,NOTE.get(lab,""),al=lft,font=sf)
    prev3=v3; r+=1
# 연도 요약행
ws.cell(r,1,"※ 연간(FY): 2023 ≈ 6,726억 → 2024 15,479억 → 2025 22,017억 (억원)").font=sf
ws.cell(r,1).alignment=lft; ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=ncol)
widths=[9,11,16,14,10,16,34]
for j,w in enumerate(widths,1): ws.column_dimensions[get_column_letter(j)].width=w
ws.freeze_panes="A4"

# ============ 원자료_10%고객 ============
raw=wb.create_sheet("원자료_10%고객")
cols=["보고서","구분","고객","부문","당분기(3개월)","누적/당기(YTD·연간)","각주"]
raw.merge_cells(start_row=1,start_column=1,end_row=1,end_column=len(cols))
c=raw.cell(1,1,"[원자료] 연결 매출 10% 이상 외부 고객 — 보고서별 전체 (단위: 백만원)  ※ 익명 표기(고객1·2…). 폴란드는 위 시트에서 K2와 대조해 식별.")
c.font=Font(name=FONT,bold=True,color="FFFFFF",size=11); c.fill=navy; c.alignment=lft; raw.row_dimensions[1].height=20
for j,h in enumerate(cols,1):
    cc=raw.cell(2,j,h); cc.font=hf; cc.fill=blue; cc.alignment=ctr; cc.border=bd
rr=3
for lab in ["2021Q2","2021Q3","2021Q4","2022Q1","2022Q2","2022Q3","2022Q4","2023Q1","2023Q2","2023Q3","2023Q4","2024Q1","2024Q2","2024Q3","2024Q4","2025Q1","2025Q2","2025Q3","2025Q4","2026Q1"]:
    d=C[lab]; custs=d["custs"]
    for ci,cu in enumerate(custs):
        pol = (lab in P["pol_idx"] and P["pol_idx"][lab]==ci)
        v3=cu["cur_3m"]; vy=cu["cur_cum"]
        vals=[lab if ci==0 else "", d["ptype"], cu["label"]+(" ★폴란드" if pol else ""), cu["seg"] or "",
              round(v3/1000) if v3 else None, round(vy/1000) if vy else None, ""]
        for j,v in enumerate(vals,1):
            cc=raw.cell(rr,j,v); cc.font=nf; cc.border=bd
            cc.alignment=rgt if isinstance(v,(int,float)) else (lft if j==3 else ctr)
            if j in (5,6): cc.number_format=WON
            if pol: cc.fill=defe
        rr+=1
    # 각주는 첫 행에
raw.column_dimensions["A"].width=9; raw.column_dimensions["B"].width=7; raw.column_dimensions["C"].width=18
raw.column_dimensions["D"].width=8; raw.column_dimensions["E"].width=15; raw.column_dimensions["F"].width=17; raw.column_dimensions["G"].width=20
raw.freeze_panes="A3"

wb.save(XLSX)
print("added sheets. total sheets:", wb.sheetnames)
