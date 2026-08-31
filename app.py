import io
import html
import re
from datetime import datetime, date
from collections import defaultdict

import pandas as pd
import streamlit as st


# ============================================================
# 基础设置
# ============================================================

PROJECT_ORDER = [
    'LAZADA',
    'COMONE_DIRECT',
    'COMONE_PANDAN',
    'TAOBAO',
    'PDD',
    'PDD_SPX',
    'CAINIAO_COM',
    'CAINIAO_COE'
]

PANDAN_PROJECTS = {
    'COMONE_DIRECT',
    'COMONE_PANDAN',
    'TAOBAO',
    'PDD',
    'PDD_SPX',
    'CAINIAO_COM'
}

VENDORS = [
    'EZBUY',
    'DDU',
    'YJD'
]

DISPLAY = {
    'LAZADA': 'LAZADA',
    'COMONE_DIRECT': 'COMONE 直达',
    'COMONE_PANDAN': 'COMONE PANDAN',
    'TAOBAO': 'TAOBAO',
    'PDD': 'PDD',
    'PDD_SPX': 'PDD-SPX',
    'CAINIAO_COM': 'CAINIAO-COM',
    'CAINIAO_COE': 'CAINIAO-COE'
}


# ============================================================
# 周次
# ============================================================

def week_no(d):
    return (
        1 if d.day <= 7
        else 2 if d.day <= 14
        else 3 if d.day <= 21
        else 4 if d.day <= 28
        else 5
    )


# ============================================================
# 日期
# ============================================================

def parse_date(v):

    if pd.isna(v) or v is None or str(v).strip() == '':
        return None

    if isinstance(v, (datetime, date, pd.Timestamp)):
        return pd.Timestamp(v).to_pydatetime()

    try:
        return pd.to_datetime(v).to_pydatetime()

    except Exception:

        try:
            return (
                pd.Timestamp('1899-12-30')
                + pd.to_timedelta(
                    float(str(v).strip()),
                    unit='D'
                )
            ).to_pydatetime()

        except Exception:
            return None


# ============================================================
# 项目判断
# ============================================================

def project_key(platform, ref):

    p = str(platform).strip().upper()

    r = (
        str(ref)
        if not pd.isna(ref)
        else ''
    )

    # COMONE
    if p == 'COMONE':

        if re.search(
            r'DIRECT',
            r,
            re.I
        ):
            return 'COMONE_DIRECT'

        return 'COMONE_PANDAN'

    # PDD-SPX
    if p in {
        'PDD-SPX',
        'PDD SPX',
        'PDD_SPX',
        'PDDSXP'
    }:
        return 'PDD_SPX'

    # 普通项目
    return {
        'LAZADA': 'LAZADA',
        'TAOBAO': 'TAOBAO',
        'PDD': 'PDD',
        'CAINIAO-COM': 'CAINIAO_COM',
        'CAINIAO-COE': 'CAINIAO_COE'
    }.get(p)


# ============================================================
# 读取 Overall
# ============================================================

def read_overall(file_bytes):

    return pd.read_excel(
        io.BytesIO(file_bytes),
        sheet_name='Overall',
        header=0
    )


# ============================================================
# 数据分析
# ============================================================

def analyze(df):

    required = [
        'Platform',
        'Container No.',
        'Cainiao B/L Ref/ Other Ref',
        'Gate Out Date',
        'Unstuffing Date',
        'Remarks For Container'
    ]

    missing = [
        c
        for c in required
        if c not in df.columns
    ]

    if missing:
        raise ValueError(
            'Overall 缺少列：'
            + ', '.join(missing)
        )

    # 项目每周柜号
    project_week = defaultdict(set)

    # 第三方每月柜号
    third_month = defaultdict(set)

    # 第三方每周柜号
    third_week = defaultdict(set)

    # ICA 项目每月柜号
    ica_project_month = defaultdict(set)

    # ICA 每月全部柜号
    ica_month = defaultdict(set)

    # 无法归类的 ICA
    ica_unassigned_month = defaultdict(set)

    months = set()

    for _, row in df.iterrows():

        platform = str(
            row.get(
                'Platform',
                ''
            )
        ).strip()

        container = str(
            row.get(
                'Container No.',
                ''
            )
        ).strip()

        if (
            not container
            or container.lower() == 'nan'
        ):
            continue

        # ----------------------------------------------------
        # 日期
        # ----------------------------------------------------

        d = (
            parse_date(
                row.get('Unstuffing Date')
            )
            or
            parse_date(
                row.get('Gate Out Date')
            )
        )

        if not d or d.year != 2026:
            continue

        mk = (
            f'{d.year:04d}-{d.month:02d}'
        )

        wk = week_no(d)

        months.add(mk)

        # ----------------------------------------------------
        # Remarks
        # ----------------------------------------------------

        remarks = str(
            row.get(
                'Remarks For Container',
                ''
            )
        )

        # ----------------------------------------------------
        # ICA / RED SEAL
        #
        # 只要包含：
        #
        # ICA RED SEAL
        #
        # 就算 ICA。
        #
        # 例如：
        #
        # ICA RED SEAL
        # ICA RED SEAL 10AM
        # Priority - ICA RED SEAL
        #
        # 都会被识别。
        # ----------------------------------------------------

        is_ica = bool(
            re.search(
                r'ICA\s*RED\s*SEAL',
                remarks,
                re.I
            )
        )

        if is_ica:

            # ICA 总柜号
            ica_month[mk].add(
                container
            )

        # ----------------------------------------------------
        # 项目
        # ----------------------------------------------------

        project = project_key(
            platform,
            row.get(
                'Cainiao B/L Ref/ Other Ref',
                ''
            )
        )

        # 如果项目无法识别
        if not project:

            if is_ica:

                ica_unassigned_month[
                    mk
                ].add(container)

            continue

        # ----------------------------------------------------
        # 项目周统计
        # ----------------------------------------------------

        project_week[
            (
                mk,
                wk,
                project
            )
        ].add(container)

        # ----------------------------------------------------
        # ICA 项目统计
        # ----------------------------------------------------

        if is_ica:

            ica_project_month[
                (
                    mk,
                    project
                )
            ].add(container)

        # ----------------------------------------------------
        # 第三方拆柜
        # ----------------------------------------------------

        if platform.upper() not in {
            'LAZADA',
            'CAINIAO-COE'
        }:

            vendor = next(
                (
                    v
                    for v in VENDORS
                    if re.search(
                        re.escape(v),
                        remarks,
                        re.I
                    )
                ),
                None
            )

            if vendor:

                third_month[
                    (
                        mk,
                        vendor
                    )
                ].add(container)

                third_week[
                    (
                        mk,
                        wk,
                        vendor
                    )
                ].add(container)

    return (
        project_week,
        third_month,
        third_week,
        ica_project_month,
        ica_month,
        ica_unassigned_month,
        sorted(
            months,
            reverse=True
        )
    )


# ============================================================
# HTML 报告
# ============================================================

def build_html(df):

    (
        pw,
        tm,
        tw,
        ipm,
        im,
        iu,
        months
    ) = analyze(df)

    out = []

    out.append(
        '''
<!doctype html>

<html lang="zh-CN">

<head>

<meta charset="UTF-8">

<meta
    name="viewport"
    content="width=device-width, initial-scale=1.0"
>

<title>商壹仓库月度看板</title>

<style>

body{
    margin:0;
    padding:24px;
    font-family:'Segoe UI','Microsoft YaHei',sans-serif;
    background:#f3f8fd;
    color:#1f2b34;
}

.wrap{
    max-width:1380px;
    margin:0 auto;
}

.hero,
.month{
    background:#fff;
    border:1px solid #d8e6f3;
    border-radius:18px;
    padding:18px;
    box-shadow:
        0 10px 20px rgba(54,98,145,.08);
    margin-bottom:16px;
}

.hero h1{
    margin:0 0 8px;
    font-size:30px;
}

.hero p{
    margin:0;
    color:#5e7083;
    font-size:14px;
    line-height:1.6;
}

.summary{
    display:grid;
    grid-template-columns:
        repeat(4,minmax(0,1fr));
    gap:10px;
    margin:14px 0 8px;
}

.pill{
    border-radius:12px;
    padding:10px 12px;
    border:1px solid #dbe8f4;
    background:#f8fbff;
}

.pill .label{
    font-size:12px;
    color:#5f7182;
}

.pill .value{
    font-size:24px;
    font-weight:700;
    color:#153a5b;
    margin-top:4px;
}

.pill.pandan{
    background:#eaf4ff;
    border-color:#cfe3fb;
}

.pill.total{
    background:#dcecff;
    border-color:#bad6f7;
}

.pill.third{
    background:#edf8f1;
    border-color:#cfe9d7;
}

.pill.ica{
    background:#f6faf8;
    border-color:#d6eadf;
}

.grid{
    display:grid;
    grid-template-columns:
        1.45fr .95fr;
    gap:16px;
    margin-top:12px;
}

.stack{
    display:grid;
    gap:12px;
}

.card{
    background:#fff;
    border:1px solid #dbe8f4;
    border-radius:12px;
    padding:12px;
}

table{
    width:100%;
    border-collapse:collapse;
}

th,
td{
    border:1px solid #dbe8f4;
    padding:7px 8px;
    text-align:center;
    font-size:12.5px;
}

th{
    background:#eef6ff;
}

.rowTotal td{
    font-weight:700;
    background:#f1f7fd;
}

.mono{
    white-space:nowrap;
}

.small{
    font-size:12px;
    color:#627485;
}

.h3{
    margin:0 0 6px;
    font-size:16px;
}

.coe{
    color:#1b5e20;
    font-weight:700;
}

.pandanCell{
    background:#dceeff;
    font-weight:700;
    color:#1f4e79;
}

.totalCell{
    background:#c7e3ff;
    font-weight:700;
    color:#18456b;
}

.legend{
    display:flex;
    gap:12px;
    flex-wrap:wrap;
    margin-top:8px;
    font-size:12px;
    color:#627485;
}

.sw{
    width:10px;
    height:10px;
    border-radius:999px;
    display:inline-block;
}

/* ==========================================================
   ICA 柜号
   ========================================================== */

.ica-list{
    margin-top:10px;
    border-top:1px solid #dbe8f4;
    padding-top:8px;
}

.ica-item{
    padding:7px 9px;
    margin:4px 0;
    background:#f7fbf9;
    border:1px solid #dcece3;
    border-radius:7px;
    font-family:
        'Consolas',
        'Courier New',
        monospace;
    font-size:13px;
    font-weight:600;
    color:#24583a;
    text-align:left;
}

.ica-count{
    display:inline-block;
    margin-left:6px;
    padding:2px 7px;
    border-radius:999px;
    background:#e1f1e8;
    color:#286043;
    font-size:11px;
}

.ica-warning{
    margin-top:10px;
    padding:8px 10px;
    border-radius:8px;
    background:#fff8e8;
    border:1px solid #f0dfaa;
    color:#7a5b13;
    font-size:12px;
    line-height:1.5;
}

@media (max-width:1180px){

    .grid{
        grid-template-columns:1fr;
    }

    .summary{
        grid-template-columns:
            repeat(2,minmax(0,1fr));
    }

}

@media (max-width:650px){

    body{
        padding:10px;
    }

    .summary{
        grid-template-columns:1fr;
    }

}

</style>

</head>

<body>

<div class="wrap">

'''
    )

    out.append(
        '''
<section class="hero">

<h1>
《商壹仓库清关拆柜月度看板》
</h1>

<p>

2026年数据。

所有项目若
<strong>Unstuffing Date</strong>
为空，会统一回退到
<strong>Gate Out Date</strong>
来归月归周。

第三方拆柜只识别
<strong>Remarks For Container</strong>；

COMONE 直达只识别
<strong>Cainiao B/L Ref/ Other Ref</strong>
中的 DIRECT。

ICA / RED SEAL 会根据
<strong>Remarks For Container</strong>
自动识别。

</p>

</section>

'''
    )

    # ========================================================
    # 每个月
    # ========================================================

    for mk in months:

        y, m = map(
            int,
            mk.split('-')
        )

        active = []

        for w in range(1, 6):

            s = (
                sum(
                    len(
                        pw[
                            (mk, w, p)
                        ]
                    )
                    for p in PROJECT_ORDER
                )
                +
                sum(
                    len(
                        tw[
                            (mk, w, v)
                        ]
                    )
                    for v in VENDORS
                )
            )

            if s:
                active.append(w)

        if not active:
            active = [1]

        month_pandan = 0
        month_overall = 0
        month_third = 0

        # ====================================================
        # ICA 全部柜号
        # ====================================================

        month_ica_containers = sorted(
            im.get(
                mk,
                set()
            )
        )

        month_ica = len(
            month_ica_containers
        )

        rows = []

        totals = {
            p: 0
            for p in PROJECT_ORDER
        }

        ps = []
        os = []
        ts = []

        # ====================================================
        # 每周
        # ====================================================

        for w in active:

            counts = {
                p: len(
                    pw[
                        (mk, w, p)
                    ]
                )
                for p in PROJECT_ORDER
            }

            for p in PROJECT_ORDER:
                totals[p] += counts[p]

            third = sum(
                len(
                    tw[
                        (mk, w, v)
                    ]
                )
                for v in VENDORS
            )

            pandan = max(
                0,
                sum(
                    counts[p]
                    for p in PANDAN_PROJECTS
                )
                - third
            )

            overall = sum(
                counts.values()
            )

            month_pandan += pandan
            month_overall += overall
            month_third += third

            start = 1 + (
                (w - 1) * 7
            )

            end = min(
                (
                    pd.Timestamp(
                        year=y,
                        month=m,
                        day=1
                    )
                    + pd.offsets.MonthEnd(0)
                ).day,
                start + 6
            )

            cells = ''.join(
                (
                    "<td class='coe'>"
                    f"{counts[p]}"
                    "</td>"
                )
                if p == 'CAINIAO_COE'

                else

                (
                    "<td>"
                    f"{counts[p]}"
                    "</td>"
                )

                for p in PROJECT_ORDER
            )

            rows.append(
                f'''
<tr>

<td>
第{w}周
</td>

<td class="mono">
{start} 至 {end}号
</td>

{cells}

<td class="pandanCell">
<strong>
{pandan}
</strong>
</td>

<td class="totalCell">
<strong>
{overall}
</strong>
</td>

</tr>
'''
            )

            ps.append(pandan)
            os.append(overall)
            ts.append(third)

        # ====================================================
        # 项目总计
        # ====================================================

        total_cells = ''.join(
            (
                "<td class='coe'>"
                f"{totals[p]}"
                "</td>"
            )
            if p == 'CAINIAO_COE'

            else

            (
                "<td>"
                f"{totals[p]}"
                "</td>"
            )

            for p in PROJECT_ORDER
        )

        # ====================================================
        # ICA 项目分布
        # ====================================================

        ica_project_rows = []

        for p in PROJECT_ORDER:

            c = len(
                ipm[
                    (mk, p)
                ]
            )

            if c:

                ica_project_rows.append(
                    f'''
<tr>

<td>
{DISPLAY[p]}
</td>

<td>
<strong>
{c}
</strong>
</td>

</tr>
'''
                )

        # 项目 ICA 总数
        classified_ica_count = sum(
            len(
                ipm[
                    (mk, p)
                ]
            )
            for p in PROJECT_ORDER
        )

        if not ica_project_rows:

            ica_project_rows = [
                '''
<tr>

<td colspan="2">
本月无已归类项目的 ICA / RED SEAL 记录
</td>

</tr>
'''
            ]

        else:

            ica_project_rows.append(
                f'''
<tr class="rowTotal">

<td>
合计
</td>

<td>
{classified_ica_count}
</td>

</tr>
'''
            )

        # ====================================================
        # ICA 柜号列表
        # ====================================================

        if month_ica_containers:

            ica_list_html = ''.join(
                f'''
<div class="ica-item">
{html.escape(container)}
</div>
'''
                for container
                in month_ica_containers
            )

        else:

            ica_list_html = '''
<div class="small">
本月没有 ICA / RED SEAL 柜。
</div>
'''

        # ====================================================
        # 未归类 ICA
        # ====================================================

        unassigned = sorted(
            iu.get(
                mk,
                set()
            )
        )

        if unassigned:

            warning_html = f'''
<div class="ica-warning">

<strong>
提示：
</strong>

本月有
<strong>
{len(unassigned)}
</strong>
个 ICA / RED SEAL 柜
没有成功归入现有项目分类，

但它们已经被 ICA 总数及柜号列表正确统计。

</div>
'''

        else:

            warning_html = ''

        # ====================================================
        # 输出月份
        # ====================================================

        out.append(
            f'''
<section class="month">

<h2>
{y}年{m}月
</h2>

<div class="small">

每个月按周查看清关、拆柜、
第三方及 ICA / RED SEAL 情况

</div>


<div class="summary">


<div class="pill pandan">

<div class="label">

PANDAN拆柜合计
PANDAN UNSTUFFING TOTAL

</div>

<div class="value">

{month_pandan}

</div>

</div>


<div class="pill total">

<div class="label">

总柜量合计
OVERALL TOTAL

</div>

<div class="value">

{month_overall}

</div>

</div>


<div class="pill third">

<div class="label">

第三方拆柜合计
THIRD-PARTY UNSTUFFING TOTAL

</div>

<div class="value">

{month_third}

</div>

</div>


<div class="pill ica">

<div class="label">

ICA / RED SEAL 合计

</div>

<div class="value">

{month_ica}

<span class="ica-count">

{month_ica} 个柜

</span>

</div>

</div>


</div>


<div class="grid">


<div class="card">


<table>

<thead>

<tr>

<th>
周次 WEEK
</th>

<th>
日期 DATE
</th>

{
    ''.join(
        (
            "<th class='coe'>"
            f"{DISPLAY[p]}"
            "</th>"
        )
        if p == 'CAINIAO_COE'

        else

        (
            "<th>"
            f"{DISPLAY[p]}"
            "</th>"
        )

        for p in PROJECT_ORDER
    )
}

<th class="pandanCell">
PANDAN拆柜
</th>

<th class="totalCell">
总柜量 OVERALL TOTAL
</th>

</tr>

</thead>


<tbody>

{''.join(rows)}


<tr class="rowTotal">

<td colspan="2">
总计
</td>

{total_cells}

<td class="pandanCell">
{month_pandan}
</td>

<td class="totalCell">
{month_overall}
</td>

</tr>


</tbody>

</table>

</div>


<div class="stack">


<!-- ======================================================
     ICA 项目分布
     ====================================================== -->

<div class="card">

<h3 class="h3">

ICA / RED SEAL 项目分布

</h3>


<table>

<thead>

<tr>

<th>
项目
</th>

<th>
柜子数
</th>

</tr>

</thead>


<tbody>

{''.join(
    ica_project_rows
)}

</tbody>

</table>

</div>


<!-- ======================================================
     ICA 柜号
     ====================================================== -->

<div class="card">

<h3 class="h3">

ICA / RED SEAL 柜号

</h3>


<div class="small">

本月共
<strong>
{month_ica}
</strong>
个 ICA / RED SEAL 柜。

每个柜号单独一行。

</div>


<div class="ica-list">

{ica_list_html}

</div>


{warning_html}

</div>


<!-- ======================================================
     趋势图
     ====================================================== -->

<div class="card">

<h3 class="h3">
趋势图
</h3>

{chart_svg(
    ps,
    os,
    ts,
    [f'第{w}周' for w in active]
)}

<div class="legend">

<span>

<i
    class="sw"
    style="background:#1f6f96"
></i>

PANDAN拆柜

</span>


<span>

<i
    class="sw"
    style="background:#4f86c6"
></i>

总柜量
OVERALL TOTAL

</span>


<span>

<i
    class="sw"
    style="background:#2d7a4c"
></i>

第三方

</span>

</div>

</div>


</div>

</div>

</section>
'''
        )

    out.append(
        '''
</div>

</body>

</html>
'''
    )

    return ''.join(out)


# ============================================================
# 趋势图
# ============================================================

def chart_svg(
    pandan,
    overall,
    third,
    labels
):

    left = 38
    right = 748
    top = 16
    bottom = 172

    n = max(
        len(labels),
        1
    )

    step = (
        0
        if n <= 1
        else (right - left) / (n - 1)
    )

    mx = max(
        1,
        max(
            pandan + overall + third,
            default=0
        )
    )

    parts = [
        "<svg viewBox='0 0 760 220'>"
    ]

    # --------------------------------------------------------
    # 网格
    # --------------------------------------------------------

    for i in range(6):

        r = i / 5

        y = (
            bottom
            - (bottom - top) * r
        )

        lab = round(
            mx * r
        )

        parts.append(
            f"""
<line
    x1='{left}'
    y1='{y}'
    x2='{right}'
    y2='{y}'
    stroke='#d9e7f3'
/>

<text
    x='30'
    y='{y+4}'
    font-size='10'
    text-anchor='end'
    fill='#667784'
>
{lab}
</text>
"""
        )

    # --------------------------------------------------------
    # 数据线
    # --------------------------------------------------------

    for vals, color in [
        (
            pandan,
            '#1f6f96'
        ),
        (
            overall,
            '#4f86c6'
        ),
        (
            third,
            '#2d7a4c'
        )
    ]:

        pts = []

        for i, v in enumerate(vals):

            x = (
                (left + right) / 2
                if n <= 1
                else left + i * step
            )

            y = (
                bottom
                - (v / mx)
                * (bottom - top)
            )

            pts.append(
                f'{x:.3f},{y:.3f}'
            )

            parts.append(
                f"""
<circle
    cx='{x:.3f}'
    cy='{y:.3f}'
    r='3'
    fill='{color}'
/>
"""
            )

        if pts:

            parts.append(
                f"""
<polyline
    points='{' '.join(pts)}'
    fill='none'
    stroke='{color}'
    stroke-width='3'
    stroke-linecap='round'
    stroke-linejoin='round'
/>
"""
            )

    # --------------------------------------------------------
    # X 轴
    # --------------------------------------------------------

    for i, label in enumerate(labels):

        x = (
            (left + right) / 2
            if n <= 1
            else left + i * step
        )

        parts.append(
            f"""
<text
    x='{x:.3f}'
    y='188'
    font-size='10'
    text-anchor='middle'
    fill='#667784'
>
{html.escape(label)}
</text>
"""
        )

    parts.append(
        '</svg>'
    )

    return ''.join(parts)


# ============================================================
# Streamlit 页面
# ============================================================

st.set_page_config(
    page_title='Cargo Arrival Report',
    page_icon='📦',
    layout='wide'
)

st.title(
    '📦 商壹仓库清关拆柜月度报告'
)

st.caption(
    '云端版：上传 Cargo Arrival (SEA & ROAD).xlsx，'
    '自动生成拆柜月度报告，并列出 ICA / RED SEAL 柜号。'
)

file = st.file_uploader(
    '上传 Excel 文件',
    type=['xlsx']
)


if file:

    if st.button(
        '生成拆柜报告',
        type='primary'
    ):

        try:

            # ------------------------------------------------
            # 读取
            # ------------------------------------------------

            df = read_overall(
                file.getvalue()
            )

            # ------------------------------------------------
            # 分析
            # ------------------------------------------------

            result = analyze(df)

            months = result[-1]

            # ------------------------------------------------
            # 生成
            # ------------------------------------------------

            report = build_html(df)

            st.success(
                f'生成成功：'
                f'共发现 {len(months)} 个月的 2026 数据。'
            )

            # ------------------------------------------------
            # 显示
            # ------------------------------------------------

            st.components.v1.html(
                report,
                height=1100,
                scrolling=True
            )

            # ------------------------------------------------
            # 下载
            # ------------------------------------------------

            st.download_button(
                '⬇️ 下载 HTML 报告',
                data=report.encode('utf-8'),
                file_name='商壹仓库_月度周报.html',
                mime='text/html'
            )

        except Exception as e:

            st.error(
                f'处理失败：{e}'
            )
