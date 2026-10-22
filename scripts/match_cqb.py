# -*- coding: utf-8 -*-
"""重庆银行产品匹配：精确代码 + 模糊名称（统一全半角括号、去份额后缀）"""
import pandas as pd
import re

DATA1 = r'C:\Users\Administrator\Downloads\银行理财产品查询 (8).xlsx'
DATA2 = r'C:\Users\Administrator\Downloads\银行理财产品查询 (10).xlsx'
NEW_LIST = r'D:\xwechat_files\wxid_dvo0b2wrw65x22_105f\temp\RWTemp\2026-10\92d32b52f7a24bcef41129f604b577ac\代销理财当前在售清单（2026.10.15）xlsx.xlsx'

print('读全量表...')
df = pd.concat([pd.read_excel(DATA1, header=1), pd.read_excel(DATA2, header=1)], ignore_index=True)
# 过滤产品名称为空的脏记录，避免三级token匹配时空集合误命中
df = df.dropna(subset=['产品名称']).reset_index(drop=True)
print(f'全量表(过滤空名称后): {len(df)}')

new = pd.read_excel(NEW_LIST, header=1)
# 新清单列名带换行符，按位置重命名前8列
new = new.iloc[:, :8]
new.columns = ['管理机构','产品代码','产品名称','产品类型','产品特点','风险等级','产品封闭期','销售时限']
new = new.dropna(subset=['产品代码']).reset_index(drop=True)
print(f'新表产品数: {len(new)}')

def norm_name(s):
    """标准化名称：全角转半角括号，去掉份额后缀，去空格"""
    if pd.isna(s): return ''
    s = str(s).strip()
    s = s.replace('（','(').replace('）',')')
    s = re.sub(r'[-—\-\s]*[A-Z](份额)?$', '', s)  # 去掉"-B份额"或"-B"结尾
    s = re.sub(r'\s+', '', s)
    return s

def extract_tokens(s):
    """提取名称中的关键token：产品号(XX号)、期限(XX天/XX日)、理财子前缀+系列"""
    s = norm_name(s)
    # 提取产品号：数字+号
    nums_hao = re.findall(r'(\d+)\s*号', s)
    # 提取期限：数字+天/日/个月
    terms = re.findall(r'(\d+)\s*(天|日|个月|月)', s)
    terms_nums = [t[0] for t in terms]
    # 去掉所有数字后的核心前缀（理财子+系列名）
    core = re.sub(r'\d+\s*号', '', s)
    core = re.sub(r'\d+\s*(天|日|个月|月)', '', core)
    core = re.sub(r'\(.*?\)', '', core)
    core = re.sub(r'\(\)', '', core)
    return set(nums_hao), set(terms_nums), core

df['code_s'] = df['银行产品编码'].astype(str).str.strip()
df['name_n'] = df['产品名称'].apply(norm_name)

matched_rows = []
missing = []
for _, r in new.iterrows():
    code = str(r['产品代码']).strip()
    name = str(r['产品名称']).strip()
    name_n = norm_name(name)
    # 1. 先按代码
    hit = df[df['code_s']==code]
    # 2. 代码没匹配，用标准化名称精确匹配
    if not len(hit):
        hit = df[df['name_n']==name_n]
    # 3. 再放宽：提取产品号集合+期限集合+核心前缀，交叉匹配（号数和期限位置互换也能匹配）
    if not len(hit):
        new_hao, new_term, new_core = extract_tokens(name)
        def _match(row):
            rh, rt, rc = extract_tokens(row['产品名称'])
            # 核心前缀必须包含
            if new_core not in rc and rc not in new_core: return False
            # 产品号集合有交集
            if new_hao and not (new_hao & rh): return False
            # 期限集合有交集
            if new_term and not (new_term & rt): return False
            return True
        cand = df[df.apply(_match, axis=1)]
        if len(cand):
            hit = cand
    if len(hit):
        row = hit.iloc[0].copy()
        row['销售机构'] = '重庆银行'
        matched_rows.append(row)
    else:
        missing.append((code, name))
        print(f'  MISS: {name[:50]}')

matched_df = pd.DataFrame(matched_rows)
# 双保险：过滤产品名称为空的行
matched_df = matched_df.dropna(subset=['产品名称']).reset_index(drop=True)
print(f'\n匹配到: {len(matched_df)}')
print(f'缺失: {len(missing)}')
print(f'缺失率: {len(missing)/len(new)*100:.1f}%')
matched_df.to_pickle(r'C:\Users\Administrator\WorkBuddy\2026-10-16-09-11-37\repo\scripts\_cqb_matched.pkl')
print('已保存')
