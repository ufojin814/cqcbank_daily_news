# -*- coding: utf-8 -*-
"""方案二：三张分表"""
import pandas as pd
import numpy as np
import re, os
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

TARGET_BANKS = ['重庆农村商业银行','重庆三峡银行','重庆富民银行','重庆银行','成都银行','汉口银行','厦门银行','哈尔滨银行','广东南粤银行','大连银行']
CQB = '重庆银行'
REPORT_DIR = r'D:\理财信息周报项目\reports'
DATA1 = r'C:\Users\Administrator\Downloads\银行理财产品查询 (6).xlsx'
DATA2 = r'C:\Users\Administrator\Downloads\银行理财产品查询 (7).xlsx'
BENCH_COL = '业绩比较基准（2026-10-01至2026-10-08）'
FIN_SUBS = ['宁银理财','招银理财','农银理财','杭银理财','华夏理财','北银理财','光大理财']
LOCAL_BANKS = ['重庆三峡银行','重庆农村商业银行','重庆富民银行']
OTHER_BANKS = ['成都银行','汉口银行','厦门银行','哈尔滨银行','广东南粤银行','大连银行']

def safe_float(v):
    try:
        if pd.isna(v): return None
        return float(v)
    except: return None
def get_main_bank(s):
    if pd.isna(s): return ''
    s=str(s)
    # 别名映射：代销机构字段里的写法 → 标准机构名
    alias = {'厦门国际银行':'厦门银行'}
    for k,v in alias.items():
        if k in s: return v
    for b in TARGET_BANKS:
        if b in s: return b
    return '其他'

def get_all_target_banks(s, exclude=None):
    """返回代销机构字段中包含的所有目标银行列表"""
    if pd.isna(s): return []
    s=str(s)
    alias = {'厦门国际银行':'厦门银行'}
    found=[]
    for k,v in alias.items():
        if k in s and v not in found: found.append(v)
    for b in TARGET_BANKS:
        if b in s and b not in found: found.append(b)
    if exclude:
        found=[b for b in found if b not in exclude]
    return found
def parse_benchmark(text):
    if pd.isna(text) or str(text).strip()=='': return None,None
    s=str(text).strip()
    hook=['指数','利率','挂钩','×','＊','*','活期','存款','收益率','中债','国债','信用','利率债','信用债','股票','基金','黄金','原油','商品','汇率','美元','欧元','港元','＋','+','组合','篮子','定存','定期','LPR','shibor','Shibor','SHIBOR','LIBOR']
    if any(k in s for k in hook): return None,None
    nums=re.findall(r'(\d+\.?\d*)\s*%',s)
    if not nums: return None,None
    nums=[float(n) for n in nums]
    if len(nums)==1: return nums[0],nums[0]
    return (min(nums)+max(nums))/2,min(nums)
def get_return(row):
    for col,label in [('近1年年化收益率（%）','近1年'),('近6月年化收益率（%）','近6月'),('近3月年化收益率（%）','近3月'),('近1月年化收益率（%）','近1月')]:
        v=safe_float(row.get(col))
        if v is not None: return v,label
    return None,''
def get_dd(row):
    for col,label in [('近1年最大回撤（%）','近1年'),('近6月最大回撤（%）','近6月'),('近3月最大回撤（%）','近3月'),('近1月最大回撤（%）','近1月')]:
        v=safe_float(row.get(col))
        if v is not None: return v,label
    return None,''

print('读数据...')
df1=pd.read_excel(DATA1,header=1); df2=pd.read_excel(DATA2,header=1)
df=pd.concat([df1,df2],ignore_index=True)
# 过滤产品名称为空的脏记录
df = df.dropna(subset=['产品名称']).reset_index(drop=True)
df['销售机构']=df['代销机构'].apply(get_main_bank)
df=df[df['销售机构'].isin(TARGET_BANKS)].copy()

# 重庆银行替换为新清单匹配到的
cqb_matched = pd.read_pickle(r'D:\理财信息周报项目\scripts\_cqb_matched.pkl')
cqb_matched['销售机构']=CQB
df = pd.concat([df[df['销售机构']!=CQB], cqb_matched], ignore_index=True)
print(f'总表: {len(df)}')

df['mid'],df['low'] = zip(*df[BENCH_COL].apply(parse_benchmark))
df['is_num']=df['mid'].notna()

# 重庆银行各类型基准（最高2只代表的最高中枢/下限）
cqb_ref={}
for st in df['运作子模式'].dropna().unique():
    cn=df[(df['运作子模式']==st)&(df['销售机构']==CQB)&(df['mid'].notna())]
    if len(cn):
        top2=cn.sort_values('mid',ascending=False).head(2)
        cqb_ref[st]=(top2['mid'].max(),top2['low'].max())
print('重庆银行基准:',cqb_ref)

now=datetime.now().strftime('%Y%m%d_%H%M')

# 样式
red_fill=PatternFill('solid',fgColor='fecaca')
red_font=Font(bold=True,color='991b1b',name='微软雅黑')
green_fill=PatternFill('solid',fgColor='dcfce7')
green_font=Font(bold=True,name='微软雅黑')
hdr_font=Font(bold=True,color='FFFFFF',size=12,name='微软雅黑')
hdr_fill=PatternFill('solid',fgColor='1e3a8a')
normal_font=Font(name='微软雅黑',size=11)
center=Alignment(horizontal='center',vertical='center',wrap_text=True)

HEADERS=['代销机构','发行机构','产品名称','业绩基准中枢(%)','业绩基准下限(%)','近1年收益(%)','近1年回撤(%)','产品类型']

def write_sheet(ws, rows_df, title):
    ws.title=title
    for c,h in enumerate(HEADERS,1):
        cell=ws.cell(1,c,h); cell.font=hdr_font; cell.fill=hdr_fill; cell.alignment=center
    r=2
    for _,row in rows_df.iterrows():
        ws.cell(r,1,row['销售机构'])
        ws.cell(r,2,row.get('发行机构',''))
        ws.cell(r,3,row.get('产品名称',''))
        if pd.notna(row.get('mid')):
            ws.cell(r,4,round(row['mid'],2)); ws.cell(r,5,round(row['low'],2))
        else:
            t=str(row.get(BENCH_COL,'无') or '无')
            ws.cell(r,4,t); ws.cell(r,5,t)
        rv,rl=get_return(row); dv,dl=get_dd(row)
        ws.cell(r,6, f'{round(rv,2)}({rl})' if rv and rl!='近1年' else (round(rv,2) if rv else '无'))
        ws.cell(r,7, f'{round(dv,2)}({dl})' if dv and dl!='近1年' else (round(dv,2) if dv else '无'))
        ws.cell(r,8,row.get('运作子模式',''))
        # 标红：同类型双高
        st=row.get('运作子模式','')
        ref=cqb_ref.get(st)
        if ref and pd.notna(row.get('mid')) and row['mid']>ref[0] and row['low']>ref[1]:
            for col in range(1,9): ws.cell(r,col).fill=red_fill; ws.cell(r,col).font=red_font
        elif row['销售机构']==CQB:
            for col in range(1,9): ws.cell(r,col).fill=green_fill; ws.cell(r,col).font=green_font
        else:
            for col in range(1,9): ws.cell(r,col).font=normal_font
        r+=1
    ws.freeze_panes='A2'
    widths=[16,16,40,14,14,14,14,14]
    from openpyxl.utils import get_column_letter
    for i,w in enumerate(widths,1): ws.column_dimensions[get_column_letter(i)].width=w

# 表1：重庆银行产品
cqb_rows = df[df['销售机构']==CQB].sort_values('mid',ascending=False)
print(f'表1重庆银行: {len(cqb_rows)}只')

# 母份额筛选（表2表3用，表1不用）：分级份额类型为空或"总份额"（即母份额）
mufen_mask = df['分级份额类型'].isna() | (df['分级份额类型'].astype(str).str.strip()=='总份额')
df_mufen = df[mufen_mask].copy()

# 兜底：没有母份额只有子份额的产品，取A类或第一个子份额作为代表
def get_base_code(code):
    """去掉银行产品编码末尾单个大写字母（份额标识），得到产品基础代码"""
    if pd.isna(code): return ''
    s=str(code).strip()
    return re.sub(r'[A-Z]$', '', s)

mufen_base_codes = set(df_mufen['银行产品编码'].apply(get_base_code))
sub_df = df[~mufen_mask & (df['分级份额类型'].astype(str).str.strip()=='子份额')].copy()
sub_df['_base'] = sub_df['银行产品编码'].apply(get_base_code)
sub_no_mufen = sub_df[~sub_df['_base'].isin(mufen_base_codes)].copy()
# 按基础代码分组，优先取A类（代码末尾为A），否则取第一个
backup_rows=[]
for base, grp in sub_no_mufen.groupby('_base'):
    a_rows = grp[grp['银行产品编码'].astype(str).str.strip().str.endswith('A')]
    if len(a_rows):
        backup_rows.append(a_rows.iloc[0])
    else:
        backup_rows.append(grp.iloc[0])
if backup_rows:
    df_backup = pd.DataFrame(backup_rows).drop(columns=['_base'], errors='ignore')
    df_mufen = pd.concat([df_mufen, df_backup], ignore_index=True)
print(f'母份额: {len(df[mufen_mask])}只, 兜底子份额: {len(backup_rows)}只, 合计表2表3用: {len(df_mufen)}只')

def expand_by_bank(data, bank_list, exclude=CQB):
    """按代销机构展开：一只产品被多家银行代销时，每家银行各占一行"""
    expanded=[]
    for _, row in data.iterrows():
        banks = get_all_target_banks(row.get('代销机构',''), exclude=exclude)
        banks = [b for b in banks if b in bank_list]
        for b in banks:
            r = row.copy()
            r['销售机构'] = b
            expanded.append(r)
    if not expanded: return pd.DataFrame(columns=data.columns)
    return pd.DataFrame(expanded)

# 表2：重庆本土银行（三峡银行、重庆农商行、富民银行）自营及代销，排除重庆银行，仅母份额，按代销机构展开
loc_rows = expand_by_bank(df_mufen, LOCAL_BANKS).sort_values(['销售机构','mid'],ascending=[True,False])
print(f'表2本土银行: {len(loc_rows)}只')

# 表3：其他在渝销售城商行（成都、汉口、厦门、哈尔滨、南粤、大连）代销，排除重庆银行，仅母份额，按代销机构展开
oth_rows = expand_by_bank(df_mufen, OTHER_BANKS).sort_values(['销售机构','发行机构','mid'],ascending=[True,True,False])
print(f'表3其他城商行: {len(oth_rows)}只')

# 生成Excel
wb=Workbook()
ws1=wb.active; write_sheet(ws1, cqb_rows, '重庆银行产品')
ws2=wb.create_sheet(); write_sheet(ws2, loc_rows, '重庆本土银行产品')
ws3=wb.create_sheet(); write_sheet(ws3, oth_rows, '其他城商行代销产品')

out=os.path.join(REPORT_DIR,f'重庆地区代销理财对比_{now}.xlsx')
wb.save(out)
print(f'\n方案二Excel: {out}')
print(f'表1: {len(cqb_rows)}只 | 表2: {len(loc_rows)}只 | 表3: {len(oth_rows)}只')

# 同时输出产品总表（全量字段）
from openpyxl import Workbook as WB
from openpyxl.utils import get_column_letter as gcl
wb_t=WB(); ws_t=wb_t.active; ws_t.title='产品总表'
ac=list(df.columns)
for c,col in enumerate(ac,1):
    cell=ws_t.cell(1,c,col); cell.font=Font(bold=True,color='FFFFFF',name='微软雅黑'); cell.fill=PatternFill('solid',fgColor='1e3a8a')
for ri,row in enumerate(df.itertuples(index=False),2):
    for ci,val in enumerate(row,1):
        ws_t.cell(ri,ci,val if pd.notna(val) else '')
ws_t.freeze_panes='A2'; ws_t.auto_filter.ref=f'A1:{gcl(len(ac))}{len(df)+1}'
tp=os.path.join(REPORT_DIR,f'重庆地区城商行理财_产品总表_{now}.xlsx')
wb_t.save(tp); print(f'产品总表: {tp}')
