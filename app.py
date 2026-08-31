import io, html, re
from datetime import datetime, date
from collections import defaultdict
import pandas as pd
import streamlit as st

PROJECT_ORDER = ['LAZADA','COMONE_DIRECT','COMONE_PANDAN','TAOBAO','PDD','CAINIAO_COM','CAINIAO_COE']
PANDAN_PROJECTS = {'COMONE_DIRECT','COMONE_PANDAN','TAOBAO','PDD','CAINIAO_COM'}
VENDORS = ['EZBUY','DDU','YJD']
DISPLAY = {
    'LAZADA':'LAZADA','COMONE_DIRECT':'COMONE 直达','COMONE_PANDAN':'COMONE PANDAN',
    'TAOBAO':'TAOBAO','PDD':'PDD','CAINIAO_COM':'CAINIAO-COM','CAINIAO_COE':'CAINIAO-COE'
}

def week_no(d):
    return 1 if d.day <= 7 else 2 if d.day <= 14 else 3 if d.day <= 21 else 4 if d.day <= 28 else 5

def parse_date(v):
    if pd.isna(v) or v is None or str(v).strip() == '': return None
    if isinstance(v, (datetime, date, pd.Timestamp)):
        return pd.Timestamp(v).to_pydatetime()
    try:
        return pd.to_datetime(v).to_pydatetime()
    except Exception:
        try:
            return pd.Timestamp('1899-12-30') + pd.to_timedelta(float(str(v).strip()), unit='D')
        except Exception:
            return None

def project_key(platform, ref):
    p = str(platform).strip().upper()
    r = str(ref) if not pd.isna(ref) else ''
    if p == 'COMONE': return 'COMONE_DIRECT' if re.search('DIRECT', r, re.I) else 'COMONE_PANDAN'
    return {'LAZADA':'LAZADA','TAOBAO':'TAOBAO','PDD':'PDD','CAINIAO-COM':'CAINIAO_COM','CAINIAO-COE':'CAINIAO_COE'}.get(p)

def read_overall(file_bytes):
    return pd.read_excel(io.BytesIO(file_bytes), sheet_name='Overall', header=0)

def analyze(df):
    required = ['Platform','Container No.','Cainiao B/L Ref/ Other Ref','Gate Out Date','Unstuffing Date','Remarks For Container']
    missing = [c for c in required if c not in df.columns]
    if missing: raise ValueError('Overall 缺少列：' + ', '.join(missing))
    project_week = defaultdict(set)
    third_month = defaultdict(set)
    third_week = defaultdict(set)
    ica_month = defaultdict(set)
    months = set()
    for _, row in df.iterrows():
        platform = str(row.get('Platform','')).strip()
        container = str(row.get('Container No.','')).strip()
        if not platform or not container or container.lower() == 'nan': continue
        project = project_key(platform, row.get('Cainiao B/L Ref/ Other Ref',''))
        if not project: continue
        d = parse_date(row.get('Unstuffing Date')) or parse_date(row.get('Gate Out Date'))
        if not d or d.year != 2026: continue
        mk = f'{d.year:04d}-{d.month:02d}'; wk = week_no(d)
        months.add(mk)
        project_week[(mk,wk,project)].add(container)
        remarks = str(row.get('Remarks For Container',''))
        if re.search(r'ICA|RED\s*SEAL', remarks, re.I): ica_month[(mk,project)].add(container)
        if platform.upper() not in {'LAZADA','CAINIAO-COE'}:
            vendor = next((v for v in VENDORS if re.search(v, remarks, re.I)), None)
            if vendor:
                third_month[(mk,vendor)].add(container)
                third_week[(mk,wk,vendor)].add(container)
    return project_week, third_month, third_week, ica_month, sorted(months, reverse=True)

def build_html(df):
    pw, tm, tw, im, months = analyze(df)
    out=[]
    out.append('''<!doctype html><html lang="zh-CN"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>商壹仓库月度看板</title>
<style>body{margin:0;padding:24px;font-family:'Segoe UI','Microsoft YaHei',sans-serif;background:#f3f8fd;color:#1f2b34}.wrap{max-width:1380px;margin:0 auto}.hero,.month{background:#fff;border:1px solid #d8e6f3;border-radius:18px;padding:18px;box-shadow:0 10px 20px rgba(54,98,145,.08);margin-bottom:16px}.hero h1{margin:0 0 8px;font-size:30px}.hero p{margin:0;color:#5e7083;font-size:14px;line-height:1.6}.summary{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:10px;margin:14px 0 8px}.pill{border-radius:12px;padding:10px 12px;border:1px solid #dbe8f4;background:#f8fbff}.pill .label{font-size:12px;color:#5f7182}.pill .value{font-size:24px;font-weight:700;color:#153a5b;margin-top:4px}.pill.pandan{background:#eaf4ff;border-color:#cfe3fb}.pill.total{background:#dcecff;border-color:#bad6f7}.pill.third{background:#edf8f1;border-color:#cfe9d7}.pill.ica{background:#f6faf8;border-color:#d6eadf}.grid{display:grid;grid-template-columns:1.45fr .95fr;gap:16px;margin-top:12px}.stack{display:grid;gap:12px}.card{background:#fff;border:1px solid #dbe8f4;border-radius:12px;padding:12px}table{width:100%;border-collapse:collapse}th,td{border:1px solid #dbe8f4;padding:7px 8px;text-align:center;font-size:12.5px}th{background:#eef6ff}.rowTotal td{font-weight:700;background:#f1f7fd}.mono{white-space:nowrap}.small{font-size:12px;color:#627485}.h3{margin:0 0 6px;font-size:16px}.coe{color:#1b5e20;font-weight:700}.pandanCell{background:#dceeff;font-weight:700;color:#1f4e79}.totalCell{background:#c7e3ff;font-weight:700;color:#18456b}.legend{display:flex;gap:12px;flex-wrap:wrap;margin-top:8px;font-size:12px;color:#627485}.sw{width:10px;height:10px;border-radius:999px;display:inline-block}@media (max-width:1180px){.grid{grid-template-columns:1fr}.summary{grid-template-columns:repeat(2,minmax(0,1fr))}}</style></head><body><div class="wrap">''')
    out.append('<section class="hero"><h1>《商壹仓库清关拆柜月度看板》</h1><p>2026年数据。所有项目若 <strong>Unstuffing Date</strong> 为空，会统一回退到 <strong>Gate Out Date</strong> 来归月归周。第三方拆柜只识别 <strong>Remarks For Container</strong>；COMONE 直达只识别 <strong>Cainiao B/L Ref/ Other Ref</strong> 中的 DIRECT。</p></section>')
    for mk in months:
        y,m=map(int,mk.split('-'))
        active=[]
        for w in range(1,6):
            s=sum(len(pw[(mk,w,p)]) for p in PROJECT_ORDER)+sum(len(tw[(mk,w,v)]) for v in VENDORS)
            if s: active.append(w)
        if not active: active=[1]
        month_pandan=month_overall=month_third=0
        month_ica=sum(len(im[(mk,p)]) for p in PROJECT_ORDER)
        rows=[]; totals={p:0 for p in PROJECT_ORDER}
        ps=[]; os=[]; ts=[]
        for w in active:
            counts={p:len(pw[(mk,w,p)]) for p in PROJECT_ORDER}
            for p in PROJECT_ORDER: totals[p]+=counts[p]
            third=sum(len(tw[(mk,w,v)]) for v in VENDORS)
            pandan=max(0,sum(counts[p] for p in PANDAN_PROJECTS)-third)
            overall=sum(counts.values())
            month_pandan+=pandan; month_overall+=overall; month_third+=third
            start=1+(w-1)*7; end=min((pd.Timestamp(year=y,month=m,day=1)+pd.offsets.MonthEnd(0)).day,start+6)
            cells=''.join("<td class='coe'>%s</td>" % counts[p] if p=='CAINIAO_COE' else "<td>%s</td>" % counts[p] for p in PROJECT_ORDER)
            rows.append(f"<tr><td>第{w}周</td><td class='mono'>{start} 至 {end}号</td>{cells}<td class='pandanCell'><strong>{pandan}</strong></td><td class='totalCell'><strong>{overall}</strong></td></tr>")
            ps.append(pandan); os.append(overall); ts.append(third)
        total_cells=''.join("<td class='coe'>%s</td>" % totals[p] if p=='CAINIAO_COE' else "<td>%s</td>" % totals[p] for p in PROJECT_ORDER)
        ica_rows=[]
        for p in PROJECT_ORDER:
            c=len(im[(mk,p)])
            if c: ica_rows.append(f"<tr><td>{DISPLAY[p]}</td><td><strong>{c}</strong></td></tr>")
        if not ica_rows: ica_rows=['<tr><td colspan="2">本月无 ICA / RED SEAL 记录</td></tr>']
        else: ica_rows.append(f'<tr class="rowTotal"><td>合计</td><td>{month_ica}</td></tr>')
        out.append(f'''<section class="month"><h2>{y}年{m}月</h2><div class="small">每个月按周查看清关、拆柜、第三方及 ICA / RED SEAL 情况</div><div class="summary"><div class="pill pandan"><div class="label">PANDAN拆柜合计 PANDAN UNSTUFFING TOTAL</div><div class="value">{month_pandan}</div></div><div class="pill total"><div class="label">总柜量合计 OVERALL TOTAL</div><div class="value">{month_overall}</div></div><div class="pill third"><div class="label">第三方拆柜合计 THIRD-PARTY UNSTUFFING TOTAL</div><div class="value">{month_third}</div></div><div class="pill ica"><div class="label">ICA / RED SEAL 合计</div><div class="value">{month_ica}</div></div></div><div class="grid"><div class="card"><table><thead><tr><th>周次 WEEK</th><th>日期 DATE</th>{''.join("<th class='coe'>%s</th>" % DISPLAY[p] if p=='CAINIAO_COE' else "<th>%s</th>" % DISPLAY[p] for p in PROJECT_ORDER)}<th class="pandanCell">PANDAN拆柜</th><th class="totalCell">总柜量 OVERALL TOTAL</th></tr></thead><tbody>{''.join(rows)}<tr class="rowTotal"><td colspan="2">总计</td>{total_cells}<td class="pandanCell">{month_pandan}</td><td class="totalCell">{month_overall}</td></tr></tbody></table></div><div class="stack"><div class="card"><h3 class="h3">ICA / RED SEAL 项目分布</h3><table><thead><tr><th>项目</th><th>柜子数</th></tr></thead><tbody>{''.join(ica_rows)}</tbody></table></div><div class="card"><h3 class="h3">趋势图</h3>{chart_svg(ps,os,ts,[f'第{w}周' for w in active])}<div class="legend"><span><i class="sw" style="background:#1f6f96"></i>PANDAN拆柜</span><span><i class="sw" style="background:#4f86c6"></i>总柜量 OVERALL TOTAL</span><span><i class="sw" style="background:#2d7a4c"></i>第三方</span></div></div></div></div></section>''')
    out.append('</div></body></html>')
    return ''.join(out)

def chart_svg(pandan, overall, third, labels):
    left,right,top,bottom=38,748,16,172; n=max(len(labels),1); step=0 if n<=1 else (right-left)/(n-1); mx=max(1,max(pandan+overall+third,default=0)); parts=["<svg viewBox='0 0 760 220'>"]
    for i in range(6):
        r=i/5; y=bottom-(bottom-top)*r; lab=round(mx*r); parts.append(f"<line x1='{left}' y1='{y}' x2='{right}' y2='{y}' stroke='#d9e7f3'/><text x='30' y='{y+4}' font-size='10' text-anchor='end' fill='#667784'>{lab}</text>")
    for vals,color in [(pandan,'#1f6f96'),(overall,'#4f86c6'),(third,'#2d7a4c')]:
        pts=[]
        for i,v in enumerate(vals):
            x=(left+right)/2 if n<=1 else left+i*step; y=bottom-(v/mx)*(bottom-top); pts.append(f'{x:.3f},{y:.3f}'); parts.append(f"<circle cx='{x:.3f}' cy='{y:.3f}' r='3' fill='{color}'/>")
        parts.insert(-len(vals), f"<polyline points='{' '.join(pts)}' fill='none' stroke='{color}' stroke-width='3' stroke-linecap='round' stroke-linejoin='round'/>")
    for i,l in enumerate(labels):
        x=(left+right)/2 if n<=1 else left+i*step; parts.append(f"<text x='{x:.3f}' y='188' font-size='10' text-anchor='middle' fill='#667784'>{html.escape(l)}</text>")
    parts.append('</svg>'); return ''.join(parts)

st.set_page_config(page_title='Cargo Arrival Report', page_icon='📦', layout='wide')
st.title('📦 商壹仓库清关拆柜月度报告')
st.caption('第一版云端测试：上传 Cargo Arrival (SEA & ROAD).xlsx，只生成拆柜月度报告。')
file = st.file_uploader('上传 Excel 文件', type=['xlsx'])
if file:
    if st.button('生成拆柜报告', type='primary'):
        try:
            df=read_overall(file.getvalue())
            pw,tm,tw,im,months=analyze(df)
            report=build_html(df)
            st.success(f'生成成功：共发现 {len(months)} 个月的 2026 数据。')
            st.components.v1.html(report, height=900, scrolling=True)
            st.download_button('⬇️ 下载 HTML 报告', data=report.encode('utf-8'), file_name='商壹仓库_月度周报.html', mime='text/html')
        except Exception as e:
            st.error(f'处理失败：{e}')
