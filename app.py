import io
import html
import re
from datetime import datetime, date
from collections import defaultdict

import pandas as pd
import streamlit as st


# ============================================================
# Project设置
# ============================================================

# CAINIAO-COE 已删除
PROJECT_ORDER = [
    'LAZADA',
    'COMONE_DIRECT',
    'COMONE_PANDAN',
    'TAOBAO',
    'PDD',
    'PDD_SPX',
    'CAINIAO_COM'
]


# 可以正常计入 PANDAN 的Project
PANDAN_PROJECTS = {
    'COMONE_PANDAN',
    'TAOBAO',
    'PDD',
    'CAINIAO_COM'
}


# 第三方拆柜识别关键词
VENDORS = [
    'EZBUY'
]


DISPLAY = {
    'LAZADA': 'LAZADA',
    'COMONE_DIRECT': 'COMONE 直达',
    'COMONE_PANDAN': 'COMONE PANDAN',
    'TAOBAO': 'TAOBAO',
    'PDD': 'PDD',
    'PDD_SPX': 'PDD-SPX',
    'CAINIAO_COM': 'CAINIAO-COM'
}


# ============================================================
# 周次
# ============================================================

def week_no(d):

    if d.day <= 7:
        return 1

    if d.day <= 14:
        return 2

    if d.day <= 21:
        return 3

    if d.day <= 28:
        return 4

    return 5


# ============================================================
# 日期处理
# ============================================================

def parse_date(v):

    if pd.isna(v) or v is None:
        return None

    if str(v).strip() == '':
        return None

    if isinstance(
        v,
        (datetime, date, pd.Timestamp)
    ):
        return pd.Timestamp(v).to_pydatetime()

    try:

        return pd.to_datetime(
            v
        ).to_pydatetime()

    except Exception:

        try:

            return (
                pd.Timestamp(
                    '1899-12-30'
                )
                +
                pd.to_timedelta(
                    float(
                        str(v).strip()
                    ),
                    unit='D'
                )
            ).to_pydatetime()

        except Exception:

            return None


# ============================================================
# Project判断
# ============================================================

def project_key(platform, ref):

    p = str(
        platform
    ).strip().upper()

    r = (
        str(ref)
        if not pd.isna(ref)
        else ''
    )


    # --------------------------------------------------------
    # COMONE
    # --------------------------------------------------------

    if p == 'COMONE':

        if re.search(
            r'DIRECT',
            r,
            re.I
        ):
            return 'COMONE_DIRECT'

        return 'COMONE_PANDAN'


    # --------------------------------------------------------
    # PDD-SPX
    #
    # 支持几种常见写法
    # --------------------------------------------------------

    if p in {
        'PDD-SPX',
        'PDD SPX',
        'PDD_SPX',
        'PDDSPX'
    }:

        return 'PDD_SPX'


    # --------------------------------------------------------
    # 普通Project
    # --------------------------------------------------------

    mapping = {

        'LAZADA':
            'LAZADA',

        'TAOBAO':
            'TAOBAO',

        'PDD':
            'PDD',

        'CAINIAO-COM':
            'CAINIAO_COM'

    }

    return mapping.get(p)


# ============================================================
# 判断是否 EZBUY 第三方
# ============================================================

def is_ezbuy(remarks):

    text = str(
        remarks
    )

    return bool(
        re.search(
            r'EZBUY',
            text,
            re.I
        )
    )


# ============================================================
# 判断是否 ICA RED SEAL
# ============================================================

def is_ica_red_seal(remarks):

    text = str(
        remarks
    )

    return bool(
        re.search(
            r'ICA\s*RED\s*SEAL',
            text,
            re.I
        )
    )


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
            'Missing required columns in the Overall sheet: '
            +
            ', '.join(missing)
        )


    # --------------------------------------------------------
    # Project / 周 / 柜号
    # --------------------------------------------------------

    project_week = defaultdict(set)


    # --------------------------------------------------------
    # 第三方 / 月 / 柜号
    # --------------------------------------------------------

    third_month = defaultdict(set)


    # --------------------------------------------------------
    # 第三方 / 周 / 柜号
    # --------------------------------------------------------

    third_week = defaultdict(set)


    # --------------------------------------------------------
    # EZBUY Third-Party / 月 / Project / 柜号
    #
    # 用于报告中列出具体 EZBUY 柜号及Project
    # --------------------------------------------------------

    third_project_month = defaultdict(set)


    # --------------------------------------------------------
    # ICA / Project / 月 / 柜号
    # --------------------------------------------------------

    ica_project_month = defaultdict(set)


    # --------------------------------------------------------
    # ICA / 月 / 全部柜号
    # --------------------------------------------------------

    ica_month = defaultdict(set)


    # --------------------------------------------------------
    # 无法归类的 ICA
    # --------------------------------------------------------

    ica_unassigned_month = defaultdict(set)


    months = set()


    # ========================================================
    # 逐行读取 Excel
    # ========================================================

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


        # 没有柜号跳过

        if (
            not container
            or
            container.lower() == 'nan'
        ):

            continue


        # ----------------------------------------------------
        # 日期
        #
        # 优先 Unstuffing Date
        # 没有则使用 Gate Out Date
        # ----------------------------------------------------

        d = (

            parse_date(
                row.get(
                    'Unstuffing Date'
                )
            )

            or

            parse_date(
                row.get(
                    'Gate Out Date'
                )
            )

        )


        # 只统计 2026

        if (
            not d
            or
            d.year != 2026
        ):

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
        # ICA
        # ----------------------------------------------------

        ica = is_ica_red_seal(
            remarks
        )


        if ica:

            # ICA 总柜号
            ica_month[
                mk
            ].add(
                container
            )


        # ----------------------------------------------------
        # Project
        # ----------------------------------------------------

        project = project_key(

            platform,

            row.get(
                'Cainiao B/L Ref/ Other Ref',
                ''
            )

        )


        # ----------------------------------------------------
        # 如果Project无法识别
        # ----------------------------------------------------

        if not project:

            if ica:

                ica_unassigned_month[
                    mk
                ].add(
                    container
                )

            continue


        # ----------------------------------------------------
        # Project每周统计
        # ----------------------------------------------------

        project_week[
            (
                mk,
                wk,
                project
            )
        ].add(
            container
        )


        # ----------------------------------------------------
        # ICA Project统计
        # ----------------------------------------------------

        if ica:

            ica_project_month[
                (
                    mk,
                    project
                )
            ].add(
                container
            )


        # ----------------------------------------------------
        # EZBUY 第三方
        #
        # 只要 Remarks For Container 包含 EZBUY，
        # 就代表这containers were handled by EZBUY for third-party unstuffing.
        #
        # EZBUY 柜：
        # 1. 计入Overall Total
        # 2. 计入原Project数量
        # 3. 计入第三方拆柜数量
        # 4. 不计入 PANDAN 拆柜数量
        # ----------------------------------------------------

        if is_ezbuy(remarks):

            third_month[
                (
                    mk,
                    'EZBUY'
                )
            ].add(
                container
            )


            third_week[
                (
                    mk,
                    wk,
                    'EZBUY'
                )
            ].add(
                container
            )


            # 保存 EZBUY 柜号 + Project
            third_project_month[
                (
                    mk,
                    project
                )
            ].add(
                container
            )


    return (

        project_week,

        third_month,

        third_week,

        third_project_month,

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

        tpm,

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

<title>Cargo Operations Dashboard</title>


<style>


body{

    margin:0;

    padding:24px;

    font-family:
        'Segoe UI',
        'Microsoft YaHei',
        sans-serif;

    background:#f3f8fd;

    color:#1f2b34;

}


.wrap{

    max-width:1450px;

    margin:0 auto;

}


.hero,
.month{

    background:#fff;

    border:1px solid #d8e6f3;

    border-radius:18px;

    padding:18px;

    box-shadow:
        0 10px 20px
        rgba(54,98,145,.08);

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

    line-height:1.7;

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
        1.5fr .9fr;

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


.ica-project{

    font-size:11px;

    color:#627485;

    margin-left:8px;

    font-family:
        'Segoe UI',
        'Microsoft YaHei',
        sans-serif;

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


.third-badge{

    display:inline-block;

    padding:2px 7px;

    margin-left:5px;

    border-radius:999px;

    background:#e7f4eb;

    color:#2d6b43;

    font-size:11px;

}


.third-list{

    margin-top:10px;

    border-top:1px solid #dbe8f4;

    padding-top:8px;

}


.third-item{

    padding:7px 9px;

    margin:4px 0;

    background:#f8fcf9;

    border:1px solid #dcece3;

    border-radius:7px;

    font-family:
        'Consolas',
        'Courier New',
        monospace;

    font-size:13px;

    font-weight:600;

    color:#285d3e;

    text-align:left;

}


.third-project{

    font-size:11px;

    color:#627485;

    margin-left:8px;

    font-family:
        'Segoe UI',
        'Microsoft YaHei',
        sans-serif;

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


    # ========================================================
    # 页面说明
    # ========================================================

    out.append(
        '''
<section class="hero">

<h1>
Cargo Operations Dashboard
</h1>

<p>

2026 data.

For all projects, if
<strong>Unstuffing Date</strong>
is unavailable, the report falls back to
<strong>Gate Out Date</strong>
for monthly and weekly classification.

<br>

<strong>PANDAN unstuffing rule:</strong>

Projects handled by the PANDAN warehouse include
COMONE PANDAN、TAOBAO、PDD、CAINIAO-COM。

LAZADA, COMONE Direct, and PDD-SPX
are excluded from PANDAN.

If the Remarks For Container field contains
<strong>EZBUY</strong>
the container is classified as third-party unstuffing.
It is still included in the overall and original project counts,
但are excluded from PANDAN.

<br>

<strong>ICA / RED SEAL:</strong>

If the Remarks For Container field contains
<strong>ICA RED SEAL</strong>
the container is automatically classified as an ICA container.

</p>

</section>

'''
    )


    # ========================================================
    # 月份
    # ========================================================

    for mk in months:

        y, m = map(
            int,
            mk.split('-')
        )


        # ----------------------------------------------------
        # 找出有数据的周
        # ----------------------------------------------------

        active = []


        for w in range(1, 6):

            project_count = sum(

                len(
                    pw[
                        (
                            mk,
                            w,
                            p
                        )
                    ]
                )

                for p in PROJECT_ORDER

            )


            third_count = len(

                tw[
                    (
                        mk,
                        w,
                        'EZBUY'
                    )
                ]

            )


            if (
                project_count
                or
                third_count
            ):

                active.append(w)


        if not active:

            active = [1]


        # ----------------------------------------------------
        # 月度变量
        # ----------------------------------------------------

        month_pandan = 0

        month_overall = 0

        month_third = len(

            tm.get(
                (
                    mk,
                    'EZBUY'
                ),
                set()
            )

        )


        # ICA 全部柜号

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
        # 每周统计
        # ====================================================

        for w in active:


            counts = {

                p: len(
                    pw[
                        (
                            mk,
                            w,
                            p
                        )
                    ]
                )

                for p in PROJECT_ORDER

            }


            # 月度累计

            for p in PROJECT_ORDER:

                totals[p] += counts[p]


            # ------------------------------------------------
            # EZBUY 第三方
            # ------------------------------------------------

            third_containers = tw[
                (
                    mk,
                    w,
                    'EZBUY'
                )
            ]


            third = len(
                third_containers
            )


            # ------------------------------------------------
            # PANDAN
            #
            # 这里按照“逐柜判断”的业务规则计算。
            #
            # 只有以下Project可以属于 PANDAN：
            #
            # COMONE PANDAN
            # TAOBAO
            # PDD
            # CAINIAO-COM
            #
            # 然后排除实际交给 EZBUY 的柜子。
            #
            # 因此：
            #
            # PDD + EZBUY
            #     → 不计 PANDAN
            #
            # PDD-SPX
            #     → 从Project层面就不属于 PANDAN
            #
            # LAZADA
            #     → 从Project层面就不属于 PANDAN
            #
            # COMONE 直达
            #     → 从Project层面就不属于 PANDAN
            # ------------------------------------------------

            pandan_containers = set()


            for p in PANDAN_PROJECTS:

                pandan_containers.update(

                    pw[
                        (
                            mk,
                            w,
                            p
                        )
                    ]

                )


            # EZBUY 柜即使属于 PDD / TAOBAO /
            # CAINIAO-COM / COMONE PANDAN，
            # 也不是 PANDAN 仓库拆柜，所以排除。

            pandan_containers -= third_containers


            pandan = len(
                pandan_containers
            )


            # ------------------------------------------------
            # Overall Total
            # ------------------------------------------------

            overall = sum(
                counts.values()
            )


            month_pandan += pandan

            month_overall += overall


            # ------------------------------------------------
            # 日期
            # ------------------------------------------------

            start = (
                1
                +
                (
                    w - 1
                ) * 7
            )


            end = min(

                (
                    pd.Timestamp(
                        year=y,
                        month=m,
                        day=1
                    )
                    +
                    pd.offsets.MonthEnd(0)
                ).day,

                start + 6

            )


            # ------------------------------------------------
            # Project单元格
            # ------------------------------------------------

            cells = ''.join(

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
Week {w}
</td>

<td class="mono">
{start} to {end}号
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


            ps.append(
                pandan
            )

            os.append(
                overall
            )

            ts.append(
                third
            )


        # ====================================================
        # Project月度Total
        # ====================================================

        total_cells = ''.join(

            (
                "<td>"
                f"{totals[p]}"
                "</td>"
            )

            for p in PROJECT_ORDER

        )


        # ====================================================
        # ICA Project分布
        # ====================================================

        ica_project_rows = []


        for p in PROJECT_ORDER:

            c = len(

                ipm[
                    (
                        mk,
                        p
                    )
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


        # 已归类 ICA 数量

        classified_ica_count = sum(

            len(
                ipm[
                    (
                        mk,
                        p
                    )
                ]
            )

            for p in PROJECT_ORDER

        )


        if not ica_project_rows:

            ica_project_rows = [

                '''
<tr>

<td colspan="2">

本月无已归类Project的
ICA / RED SEAL 记录

</td>

</tr>
'''

            ]

        else:

            ica_project_rows.append(

                f'''
<tr class="rowTotal">

<td>
Total
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

        ica_items = []


        for container in month_ica_containers:

            project_names = []


            for p in PROJECT_ORDER:

                if container in ipm[
                    (
                        mk,
                        p
                    )
                ]:

                    project_names.append(
                        DISPLAY[p]
                    )


            if project_names:

                project_text = (
                    ' / '.join(
                        project_names
                    )
                )

                project_html = (

                    f'''
<span class="ica-project">
Project：{html.escape(project_text)}
</span>
'''

                )

            else:

                project_html = ''


            ica_items.append(

                f'''
<div class="ica-item">

{html.escape(container)}

{project_html}

</div>
'''

            )


        if ica_items:

            ica_list_html = ''.join(
                ica_items
            )

        else:

            ica_list_html = '''

<div class="small">

No ICA / RED SEAL containers this month.

</div>

'''


        # ====================================================
        # EZBUY 第三方拆柜柜号列表
        # ====================================================

        third_items = []


        third_month_containers = sorted(

            tm.get(
                (
                    mk,
                    'EZBUY'
                ),
                set()
            )

        )


        for container in third_month_containers:

            project_names = []


            for p in PROJECT_ORDER:

                if container in tpm[
                    (
                        mk,
                        p
                    )
                ]:

                    project_names.append(
                        DISPLAY[p]
                    )


            if project_names:

                project_text = (
                    ' / '.join(
                        project_names
                    )
                )

                project_html = (

                    f'''
<span class="third-project">
Project：{html.escape(project_text)}
</span>
'''

                )

            else:

                project_html = ''


            third_items.append(

                f'''
<div class="third-item">

{html.escape(container)}

{project_html}

</div>
'''

            )


        if third_items:

            third_list_html = ''.join(
                third_items
            )

        else:

            third_list_html = '''

<div class="small">

No containers were handled by EZBUY this month.

</div>

'''


        # ====================================================
        # 无法归类 ICA
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
Note:
</strong>

本月有
<strong>
{len(unassigned)}
</strong>
个 ICA / RED SEAL 柜
没有成功归入现有Project分类，

but they are still included in the total ICA count and container list.

</div>

'''

        else:

            warning_html = ''


        # ====================================================
        # 月度 HTML
        # ========================================================

        out.append(

            f'''
<section class="month">

<h2>
{y}年{m}月
</h2>


<div class="small">

Review customs clearance, unstuffing, third-party handling, and ICA / RED SEAL activity by week.

</div>


<div class="summary">


<div class="pill pandan">

<div class="label">

PANDAN Unstuffing
PANDAN UNSTUFFING TOTAL

</div>

<div class="value">

{month_pandan}

</div>

</div>


<div class="pill total">

<div class="label">

Overall Containers
OVERALL TOTAL

</div>

<div class="value">

{month_overall}

</div>

</div>


<div class="pill third">

<div class="label">

Third-Party Unstuffing
THIRD-PARTY UNSTUFFING TOTAL

</div>

<div class="value">

{month_third}

</div>

</div>


<div class="pill ica">

<div class="label">

ICA / RED SEAL Total

</div>

<div class="value">

{month_ica}

<span class="ica-count">

{month_ica} containers

</span>

</div>

</div>


</div>


<div class="grid">


<!-- =====================================================
     左边：周统计
     ===================================================== -->

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

        f"<th>{DISPLAY[p]}</th>"

        for p in PROJECT_ORDER

    )
}

<th class="pandanCell">

PANDAN Unstuffing

</th>

<th class="totalCell">

Overall Total
OVERALL TOTAL

</th>

</tr>

</thead>


<tbody>

{''.join(rows)}


<tr class="rowTotal">

<td colspan="2">

Total

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


<!-- =====================================================
     右边
     ===================================================== -->

<div class="stack">


<!-- ===================================================
     ICA Project分布
     =================================================== -->

<div class="card">

<h3 class="h3">

ICA / RED SEAL by Project

</h3>


<table>

<thead>

<tr>

<th>
Project
</th>

<th>
Containers
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


<!-- ===================================================
     ICA 柜号
     =================================================== -->

<div class="card">

<h3 class="h3">

ICA / RED SEAL 柜号

</h3>


<div class="small">

This month:
<strong>
{month_ica}
</strong>
个 ICA / RED SEAL 柜。

</div>


<div class="ica-list">

{ica_list_html}

</div>


{warning_html}

</div>


<!-- ===================================================
     EZBUY 第三方拆柜柜号
     =================================================== -->

<div class="card">

<h3 class="h3">

EZBUY Third-Party Unstuffing

</h3>


<div class="small">

This month:
<strong>
{month_third}
</strong>
containers were handled by EZBUY for third-party unstuffing.

这些柜子仍计入清关Overall Total及原Project数量，
but are excluded from PANDAN unstuffing.

</div>


<div class="third-list">

{third_list_html}

</div>


</div>


<!-- ===================================================
     Weekly Trend
     =================================================== -->

<div class="card">

<h3 class="h3">

Weekly Trend

</h3>


{chart_svg(
    ps,
    os,
    ts,
    [
        f'Week {w}'
        for w in active
    ]
)}


<div class="legend">


<span>

<i
    class="sw"
    style="background:#1f6f96"
></i>

PANDAN Unstuffing

</span>


<span>

<i
    class="sw"
    style="background:#4f86c6"
></i>

Overall Total
OVERALL TOTAL

</span>


<span>

<i
    class="sw"
    style="background:#2d7a4c"
></i>

EZBUY Third-Party

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
# Weekly Trend
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

        else
        (
            right - left
        )
        /
        (
            n - 1
        )

    )


    mx = max(

        1,

        max(
            pandan
            +
            overall
            +
            third,
            default=0
        )

    )


    parts = [

        "<svg viewBox='0 0 760 220'>"

    ]


    # --------------------------------------------------------
    # 横线
    # --------------------------------------------------------

    for i in range(6):

        r = i / 5

        y = (

            bottom
            -
            (
                bottom - top
            )
            *
            r

        )


        lab = round(
            mx * r
        )


        parts.append(

            f'''
<line

x1="{left}"

y1="{y}"

x2="{right}"

y2="{y}"

stroke="#d9e7f3"

/>


<text

x="30"

y="{y+4}"

font-size="10"

text-anchor="end"

fill="#667784"

>

{lab}

</text>

'''

        )


    # --------------------------------------------------------
    # 三条线
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

                (
                    left + right
                )
                /
                2

                if n <= 1

                else

                left + i * step

            )


            y = (

                bottom
                -
                (
                    v / mx
                )
                *
                (
                    bottom - top
                )

            )


            pts.append(

                f'{x:.3f},{y:.3f}'

            )


            parts.append(

                f'''
<circle

cx="{x:.3f}"

cy="{y:.3f}"

r="3"

fill="{color}"

/>

'''

            )


        if pts:

            parts.append(

                f'''
<polyline

points="{' '.join(pts)}"

fill="none"

stroke="{color}"

stroke-width="3"

stroke-linecap="round"

stroke-linejoin="round"

/>

'''

            )


    # --------------------------------------------------------
    # X轴
    # --------------------------------------------------------

    for i, label in enumerate(labels):

        x = (

            (
                left + right
            )
            /
            2

            if n <= 1

            else

            left + i * step

        )


        parts.append(

            f'''
<text

x="{x:.3f}"

y="188"

font-size="10"

text-anchor="middle"

fill="#667784"

>

{html.escape(label)}

</text>

'''

        )


    parts.append(
        '</svg>'
    )


    return ''.join(parts)



# ============================================================
# Declarant Monthly Report
# ============================================================

DECLARANT_PROJECT_ORDER = [
    'Lazada',
    'Comone',
    'Taobao',
    'Pinduoduo',
    'Cainiao-COM',
    'Cainiao-COE',
    'Road',
    'Export',
]

DECLARANT_PROJECT_ALIASES = {
    'LAZADA': 'Lazada',
    'LAZADA.COM': 'Lazada',
    'COMONE': 'Comone',
    'COMONE_DIRECT': 'Comone',
    'COMONE_PANDAN': 'Comone',
    'TAOBAO': 'Taobao',
    'PDD': 'Pinduoduo',
    'PINDUODUO': 'Pinduoduo',
    'CAINIAO-COM': 'Cainiao-COM',
    'CAINIAO_COM': 'Cainiao-COM',
    'CAINIAO-COE': 'Cainiao-COE',
    'CAINIAO_COE': 'Cainiao-COE',
    'PDD-SPX': 'PDD-SPX',
    'PDD_SPX': 'PDD-SPX',
    'ROAD': 'Road',
    'EXPORT': 'Export',
}

DECLARANT_COLUMN_ALIASES = {
    'platform': [
        'Platform', 'Project', 'Platform Name'
    ],
    'noa_done_by': [
        'NOA Done By', 'NOA Done by', 'NOA Responsible',
        'NOA By', 'NOA负责人', 'NOA 负责人'
    ],
    'loa_done_by': [
        'LOA Done By', 'LOA Done by', 'LOA Responsible',
        'LOA By', 'LOA负责人', 'LOA 负责人'
    ],
    'done_by': [
        'Done By', 'Done by', 'Declarant', 'Declarant Name',
        'Permit Done By', '报关员', '报关员姓名'
    ],
    'noa_date': [
        'NOA Date', 'NOA date', 'NOA Date/Time',
        'NOA Completed Date', 'NOA Completion Date'
    ],
    'loa_date': [
        'LOA Date', 'LOA date', 'LOA Date/Time',
        'LOA Completed Date', 'LOA Completion Date'
    ],
    'permit_date': [
        'Permit Date', 'Permit date', 'Permit Date/Time',
        'Declaration Date', 'Declaration date',
        'Permit Completed Date', 'Done Date'
    ],
    'no_of_permits': [
        'No. of Permits', 'No of Permits', 'No. of Permit',
        'Number of Permits', 'Permit Count', 'Permits'
    ],
    'ig_number': [
        'IG number', 'IG Number', 'IG No.', 'IG No', 'IG'
    ],
    'permit_no': [
        'Permit No.', 'Permit No', 'Permit Number', 'Permit No'
    ],
}


def _find_column(df, key, required=False):
    aliases = DECLARANT_COLUMN_ALIASES[key]
    normalized = {
        re.sub(r'\s+', ' ', str(c).strip()).lower(): c
        for c in df.columns
    }
    for alias in aliases:
        hit = normalized.get(re.sub(r'\s+', ' ', alias).strip().lower())
        if hit is not None:
            return hit
    if required:
        raise ValueError(
            f"Declarant Report is missing a required column for '{key}'. "
            f"Accepted names include: {', '.join(aliases)}"
        )
    return None


def _clean_person(v):
    if pd.isna(v) or v is None:
        return ''
    s = str(v).strip()
    if not s or s.lower() in {'nan', 'none', 'nat'}:
        return ''
    return s


def _numeric(v):
    if pd.isna(v) or v is None:
        return 0
    if isinstance(v, (int, float)) and not pd.isna(v):
        return float(v)
    s = str(v).strip().replace(',', '')
    if not s:
        return 0
    try:
        return float(s)
    except Exception:
        m = re.search(r'-?\d+(?:\.\d+)?', s)
        return float(m.group()) if m else 0


def _has_value(v):
    return _clean_person(v) != ''


def _parse_project(platform):
    p = _clean_person(platform).upper().replace(' ', '')
    return DECLARANT_PROJECT_ALIASES.get(p, _clean_person(platform) or 'Unknown')


def _date_in_month(v, year=2026):
    d = parse_date(v)
    return d if d and d.year == year else None



def _month_key(v):
    d = parse_date(v)
    if not d or d.year < 2026:
        return None
    return f"{d.year:04d}-{d.month:02d}"


DECLARANT_PERSON_MAP = {
    'J': 'Jiaqian',
    'E': 'Eries',
    'M': 'Marlia',
    'JIAQIAN': 'Jiaqian',
    'ERIES': 'Eries',
    'MARLIA': 'Marlia',
}


def _declarant_name(v):
    """Map declarant values to the three standard names.

    Accept both the plain initials (E/J/M) and values such as:
    E + DRAFTED, J + DRAFT, M + DRAFTING, E/DRAFTED, etc.
    The leading E/J/M is the source of truth for assignment.
    """
    s = _clean_person(v).upper()
    if not s:
        return ''

    # Plain names / initials.
    if s in DECLARANT_PERSON_MAP:
        return DECLARANT_PERSON_MAP[s]

    # Any value whose first meaningful character is E/J/M belongs to
    # that declarant. This intentionally covers values such as:
    # E, J, M
    # E + DRAFTED, J DRAFT, M/DRAFTING
    # E DRAFTED, J + DRAFT, M anything
    # The business rule is: the leading E/J/M is the source of truth.
    m = re.match(r'^\s*([EJM])(?:\s|$|[+\-/])', s)
    if m:
        return DECLARANT_PERSON_MAP[m.group(1)]

    return ''


def _effective_overall_date(row):
    # Declarant Monthly Report: month is based on Gate Out Date.
    # Only fall back to Unstuffing Date when Gate Out Date is blank.
    d = parse_date(row.get('Gate Out Date'))
    if d:
        return d
    return parse_date(row.get('Unstuffing Date'))


def _overall_project(row):
    p = _clean_person(row.get('Platform')).upper()
    ref = _clean_person(row.get('Cainiao B/L Ref/ Other Ref'))

    if p == 'COMONE':
        return 'Comone' if re.search(r'DIRECT', ref, re.I) else 'Comone'
    if p == 'LAZADA':
        return 'Lazada'
    if p == 'TAOBAO':
        return 'Taobao'
    if p == 'PDD':
        return 'Pinduoduo'
    if p == 'PDD-SPX':
        return 'PDD-SPX'
    if p == 'CAINIAO-COM':
        return 'Cainiao-COM'
    if p == 'CAINIAO-COE':
        return 'Cainiao-COE'
    return _clean_person(row.get('Platform'))


def analyze_declarant(df, road_df=None, export_df=None):
    """
    Declarant Monthly Report based on the same Cargo workbook.

    Overall:
      - month = Gate Out Date, otherwise Unstuffing Date
      - NOA / LOA = responsible-person initials in Overall columns
      - Permit = No. of Permits, assigned by Permit/Done-By initials

    Road:
      - month = Date
      - count each IG permit record
      - assigned by Done By

    Export:
      - month = Date
      - count each Permit No. record
      - assigned by Done By
    """
    monthly = defaultdict(lambda: {
        'noa': defaultdict(int),
        'loa': defaultdict(int),
        'permit': defaultdict(float),
        'project': defaultdict(lambda: {
            'noa': defaultdict(int),
            'loa': defaultdict(int),
            'permit': defaultdict(float),
        })
    })
    people_by_month = defaultdict(set)

    def add_noa_loa(month, person, project, metric):
        if not month or not person:
            return
        monthly[month][metric][person] += 1
        monthly[month]['project'][person][metric][project] += 1
        people_by_month[month].add(person)

    def add_permit(month, person, project, value):
        if not month or not person or not value:
            return
        monthly[month]['permit'][person] += value
        monthly[month]['project'][person]['permit'][project] += value
        people_by_month[month].add(person)

    # ---------------- Overall sheet ----------------
    required_overall = [
        'Platform', 'NOA', 'LOA ', 'Permit', 'No. of Permits',
        'Unstuffing Date', 'Gate Out Date'
    ]
    missing = [c for c in required_overall if c not in df.columns]
    if missing:
        raise ValueError(
            "Overall sheet is missing required Declarant Report columns: "
            + ", ".join(missing)
        )

    for _, row in df.iterrows():
        month = _month_key(_effective_overall_date(row))
        if not month:
            continue

        project = _overall_project(row)

        noa_person = _declarant_name(row.get('NOA'))
        loa_person = _declarant_name(row.get('LOA '))
        permit_person = _declarant_name(row.get('Permit'))

        add_noa_loa(month, noa_person, project, 'noa')
        add_noa_loa(month, loa_person, project, 'loa')

        permit_count = _numeric(row.get('No. of Permits'))
        add_permit(month, permit_person, project, permit_count)

    # ---------------- Road sheet ----------------
    if road_df is not None and not road_df.empty:
        for _, row in road_df.iterrows():
            month = _month_key(row.get('Date'))
            person = _declarant_name(row.get('Done By'))
            if month and person and _has_value(row.get('Permit')):
                add_permit(month, person, 'Road', 1)

    # ---------------- Export sheet ----------------
    if export_df is not None and not export_df.empty:
        for _, row in export_df.iterrows():
            month = _month_key(row.get('Date'))
            person = _declarant_name(row.get('Done By'))
            if month and person and _has_value(row.get('Permit No.')):
                add_permit(month, person, 'Export', 1)

    # Clean integer-valued permit totals.
    for mk, data in monthly.items():
        for person, value in list(data['permit'].items()):
            if float(value).is_integer():
                data['permit'][person] = int(value)
        for person in data['project']:
            for project, value in list(data['project'][person]['permit'].items()):
                if float(value).is_integer():
                    data['project'][person]['permit'][project] = int(value)

    months = sorted(
        [m for m in monthly if m >= '2026-01'],
        reverse=True
    )

    return monthly, people_by_month, months, {}




def _project_sort_key(project):
    if project in DECLARANT_PROJECT_ORDER:
        return (0, DECLARANT_PROJECT_ORDER.index(project))
    return (1, project.lower())

def build_declarant_html(df, road_df=None, export_df=None):
    monthly, people_by_month, months, cols = analyze_declarant(df, road_df, export_df)

    if not months:
        raise ValueError("No January 2026 onward declarant activity was found.")

    out = ["""
<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Declarant Monthly Report</title>
<style>
body{margin:0;padding:24px;font-family:'Segoe UI','Microsoft YaHei',sans-serif;
background:#f3f8fd;color:#1f2b34}
.wrap{max-width:1450px;margin:0 auto}
.hero,.month{background:#fff;border:1px solid #d8e6f3;border-radius:18px;
padding:18px;box-shadow:0 10px 20px rgba(54,98,145,.08);margin-bottom:16px}
.hero h1{margin:0 0 8px;font-size:30px}
.hero p{margin:0;color:#5e7083;font-size:14px;line-height:1.7}
.cards{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:10px;margin:14px 0}
.metric{border:1px solid #dbe8f4;border-radius:12px;padding:12px;background:#f8fbff}
.metric .label{font-size:12px;color:#5f7182}
.metric .value{font-size:24px;font-weight:700;color:#153a5b;margin-top:4px}
.grid{display:grid;grid-template-columns:1.15fr .85fr;gap:16px}
.card{background:#fff;border:1px solid #dbe8f4;border-radius:12px;padding:12px}
.card h3{margin:0 0 8px;font-size:16px}
.personBlocks{display:grid;gap:12px}
.personBlock{border:1px solid #dbe8f4;border-radius:12px;background:#fbfdff;overflow:hidden}
.personTitle{padding:9px 11px;background:#eef6ff;color:#173d5f;font-weight:800;
border-bottom:1px solid #dbe8f4}
table{width:100%;border-collapse:collapse}
th,td{border:1px solid #cfddea;padding:8px 9px;text-align:center;font-size:13px}
th{background:#eaf4ff;color:#173d5f}
.total td{font-weight:800;background:#eef6ff}
.person{text-align:left;font-weight:700}
.num{font-variant-numeric:tabular-nums}
.pct{color:#48657c;min-width:85px}
.bar{height:7px;background:#e7eef6;border-radius:999px;overflow:hidden;margin-top:4px}
.bar i{display:block;height:100%;background:#2d7a4c}
.tag{display:inline-block;border-radius:999px;background:#eef6ff;border:1px solid #d4e5f6;
padding:3px 8px;font-size:12px;color:#315c7c}
.note{font-size:12px;color:#607386;margin-top:8px;line-height:1.6}
@media(max-width:980px){.cards,.grid{grid-template-columns:1fr}body{padding:14px}}
</style></head><body><div class="wrap">
<section class="hero">
<h1>Declarant Monthly Report</h1>
<p>
从 January 2026 开始统计。NOA / LOA 按对应列中的负责人记录计数；
Overall Permit 按 <strong>No. of Permits</strong> 汇总；
Road 只按 <strong>IG number</strong> 计数，Export 按 <strong>Permit No.</strong> 计数，
并都按 <strong>Done By</strong> 列归属报关员。
</p>
</section>
"""]

    for mk in months:
        y, m = map(int, mk.split('-'))
        data = monthly[mk]
        people = sorted(people_by_month[mk], key=lambda x: x.lower())

        noa_total = sum(data['noa'].values())
        loa_total = sum(data['loa'].values())
        permit_total = sum(data['permit'].values())

        out.append(f"""
<section class="month">
<h2>{datetime(y,m,1).strftime('%B %Y')}</h2>
<div class="cards">
<div class="metric"><div class="label">NOA Total</div><div class="value">{noa_total}</div></div>
<div class="metric"><div class="label">LOA Total</div><div class="value">{loa_total}</div></div>
<div class="metric"><div class="label">Permit Total</div><div class="value">{permit_total:g}</div></div>
</div>

<div class="grid">
<div class="card">
<h3>每人工作用量与占比</h3>
<table><thead><tr>
<th>报关员</th><th>NOA 数量</th><th>NOA 占比</th>
<th>LOA 数量</th><th>LOA 占比</th>
<th>Permit 数量</th><th>Permit 占比</th>
</tr></thead><tbody>
""")

        for person in people:
            noa = data['noa'].get(person, 0)
            loa = data['loa'].get(person, 0)
            permit = data['permit'].get(person, 0)
            noa_pct = noa / noa_total * 100 if noa_total else 0
            loa_pct = loa / loa_total * 100 if loa_total else 0
            permit_pct = permit / permit_total * 100 if permit_total else 0

            out.append(f"""
<tr>
<td class="person">{html.escape(person)}</td>
<td class="num">{noa}</td>
<td class="pct">{noa_pct:.1f}%<div class="bar"><i style="width:{noa_pct:.1f}%"></i></div></td>
<td class="num">{loa}</td>
<td class="pct">{loa_pct:.1f}%<div class="bar"><i style="width:{loa_pct:.1f}%"></i></div></td>
<td class="num">{permit:g}</td>
<td class="pct">{permit_pct:.1f}%<div class="bar"><i style="width:{permit_pct:.1f}%"></i></div></td>
</tr>
""")

        out.append(f"""
<tr class="total">
<td class="person">本月合计</td>
<td>{noa_total}</td><td>100.0%</td>
<td>{loa_total}</td><td>100.0%</td>
<td>{permit_total:g}</td><td>100.0%</td>
</tr>
</tbody></table>
</div>

<div class="card">
<h3>项目明细</h3>
<div class="personBlocks">
""")

        for person in people:
            pdata = data['project'][person]
            projects = set(pdata['noa']) | set(pdata['loa']) | set(pdata['permit'])
            projects = sorted(projects, key=_project_sort_key)

            out.append(f"""
<div class="personBlock">
<div class="personTitle">{html.escape(person)}</div>
<table><thead><tr><th>项目</th><th>NOA</th><th>LOA</th><th>Permit</th></tr></thead><tbody>
""")
            for project in projects:
                out.append(f"""
<tr>
<td class="person"><span class="tag">{html.escape(project)}</span></td>
<td class="num">{pdata['noa'].get(project,0)}</td>
<td class="num">{pdata['loa'].get(project,0)}</td>
<td class="num">{pdata['permit'].get(project,0):g}</td>
</tr>
""")
            out.append("</tbody></table></div>")

        out.append("""
</div>
<div class="note">
基础项目顺序：Lazada、Comone、Taobao、Pinduoduo、Cainiao-COM、Cainiao-COE、Road、Export；
新 Platform 项目会自动追加并单独显示。
</div>
</div>
</div>
</section>
""")

    out.append("""
</div></body></html>
""")
    return ''.join(out)


# ============================================================
# Streamlit Page
# ============================================================

st.set_page_config(
    page_title='Cargo Operations Dashboard',
    page_icon='📦',
    layout='wide'
)

st.markdown(
    """
    <style>
    .stApp {
        background: linear-gradient(180deg, #f8fafc 0%, #ffffff 55%);
    }
    .block-container {
        max-width: 1100px;
        padding-top: 2rem;
        padding-bottom: 4rem;
    }
    .landing {
        max-width: 760px;
        margin: 3.5rem auto 1.8rem auto;
        text-align: center;
    }
    .landing-badge {
        display: inline-block;
        padding: 0.35rem 0.8rem;
        border: 1px solid #d9e2ec;
        border-radius: 999px;
        background: rgba(255,255,255,0.85);
        color: #526579;
        font-size: 0.78rem;
        font-weight: 600;
        letter-spacing: 0.04em;
        margin-bottom: 1rem;
    }
    .landing h1 {
        font-size: 2.8rem;
        line-height: 1.15;
        letter-spacing: -0.04em;
        margin: 0 0 0.8rem 0;
        color: #17212b;
    }
    .landing p {
        max-width: 650px;
        margin: 0 auto;
        color: #687787;
        font-size: 1.03rem;
        line-height: 1.7;
    }
    [data-testid="stFileUploader"] {
        max-width: 760px;
        margin: 1.2rem auto 0 auto;
    }
    [data-testid="stFileUploaderDropzone"] {
        border: 1.5px dashed #b9c8d8 !important;
        border-radius: 16px !important;
        background: rgba(255,255,255,0.9) !important;
        padding: 1.25rem 1rem !important;
    }
    div.stButton {
        max-width: 760px;
        margin: 0.9rem auto 0 auto;
    }
    div.stButton > button {
        width: 100%;
        min-height: 2.8rem;
        border-radius: 10px;
        font-weight: 600;
        border: 0;
    }
    </style>
    <div class="landing">
        <div class="landing-badge">CARGO OPERATIONS · AUTOMATED REPORTING</div>
        <h1>📦 Cargo Operations Dashboard</h1>
        <p>
            Turn structured cargo data into clear operations and declarant reports
            in a few clicks. Upload one Excel file to get started.
        </p>
    </div>
    """,
    unsafe_allow_html=True
)

file = st.file_uploader(
    'Upload Excel File',
    type=['xlsx'],
    help='Supported format: .xlsx'
)

if file:
    if st.button('Generate Reports', type='primary'):
        try:
            file_bytes = file.getvalue()
            df = read_overall(file_bytes)
            road_df = pd.read_excel(io.BytesIO(file_bytes), sheet_name='Road')
            export_df = pd.read_excel(io.BytesIO(file_bytes), sheet_name='Export')

            # Existing Cargo Operations Report
            cargo_result = analyze(df)
            cargo_months = cargo_result[-1]
            cargo_report = build_html(df)

            # New Declarant Monthly Report
            declarant_report = build_declarant_html(df, road_df, export_df)
            declarant_result = analyze_declarant(df, road_df, export_df)
            declarant_months = declarant_result[2]

            st.success(
                f"Reports generated successfully. "
                f"Cargo: {len(cargo_months)} month(s); "
                f"Declarant: {len(declarant_months)} month(s) from Jan 2026 onward."
            )

            tab1, tab2 = st.tabs([
                '📦 Cargo Operations Report',
                '🛃 Declarant Monthly Report'
            ])

            with tab1:
                st.components.v1.html(
                    cargo_report,
                    height=1200,
                    scrolling=True
                )
                st.download_button(
                    '⬇️ Download Cargo Operations HTML',
                    data=cargo_report.encode('utf-8'),
                    file_name='Cargo_Operations_Dashboard.html',
                    mime='text/html',
                    key='download_cargo'
                )

            with tab2:
                st.components.v1.html(
                    declarant_report,
                    height=1200,
                    scrolling=True
                )
                st.download_button(
                    '⬇️ Download Declarant Monthly HTML',
                    data=declarant_report.encode('utf-8'),
                    file_name='Declarant_Monthly_Report.html',
                    mime='text/html',
                    key='download_declarant'
                )

        except Exception as e:
            st.error(f'Processing failed: {e}')
