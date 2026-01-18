# =========================================
# SOCIAL MEDIA REPORTING – FINAL PIPELINE
# =========================================

import pandas as pd
import numpy as np

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE

# -----------------------------------------
# CONFIG
# -----------------------------------------

PRIORITY_BRANDS = [
    'Similac',
    'Ensure',
    'Pediasure',
    'Pedialyte',
    'Juven',
    'Glucerna'
]

INPUT_FILE = "input.xlsx"
EXCEL_OUTPUT = "Final_Social_Media_Report.xlsx"
PPT_OUTPUT = "Final_Social_Media_Trends.pptx"

# -----------------------------------------
# BRAND ORDERING FUNCTION
# -----------------------------------------

def sort_brands(df, brand_col, volume_col):
    df = df.copy()
    df['__priority'] = df[brand_col].apply(
        lambda x: PRIORITY_BRANDS.index(x)
        if x in PRIORITY_BRANDS else len(PRIORITY_BRANDS)
    )
    df = df.sort_values(
        by=['__priority', volume_col],
        ascending=[True, False]
    )
    return df.drop(columns='__priority')

# -----------------------------------------
# LOAD & CLEAN DATA
# -----------------------------------------

df = pd.read_excel(INPUT_FILE)

df = df.rename(columns={
    'Social Channel/Platform': 'platform',
    'Brand Name/Category Name': 'brand',
    'Type of Post (Branded, Influencer,Creators, Organic, Shop)': 'post_type',
    'Post Format (Video, Reels, Shorts, Images, Carousels)': 'format',
    'Video Plays': 'views',
    'Followers': 'followers',
    'Influencer Tier': 'influencer_tier'
})

for col in ['Likes', 'Comments', 'views', 'followers']:
    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

df['engagement'] = df['Likes'] + df['Comments']

df['brand'] = df['brand'].astype(str).str.strip().str.title()

df['post_type'] = (
    df['post_type']
    .astype(str)
    .str.title()
    .replace({'Tagged': 'Influencer', 'Creators': 'Creator'})
)

df['format'] = (
    df['format']
    .astype(str)
    .str.title()
    .replace({
        'Reels': 'Shorts',
        'Images': 'Image',
        'Sidecar': 'Carousel',
        'Image+Video': 'Carousel'
    })
)

df['is_dynamic'] = df['format'].isin(['Video', 'Shorts'])
df['is_paid'] = df['post_type'].isin(['Branded', 'Influencer', 'Creator', 'Shop'])

# -----------------------------------------
# METRIC FUNCTIONS
# -----------------------------------------

def overall_metrics(x):
    video_eng = x.loc[x['is_dynamic'], 'engagement'].sum()
    video_views = x.loc[x['is_dynamic'], 'views'].sum()

    return pd.Series({
        'Total Posts': len(x),
        'Total Engagement': x['engagement'].sum(),
        'Video Engagement': video_eng,
        'Total Views': video_views,
        'Avg ER%': video_eng / video_views if video_views > 0 else None
    })

def format_metrics(x):
    fmt = x['format'].iloc[0]

    if fmt in ['Video', 'Shorts']:
        eng = x['engagement'].sum()
        views = x['views'].sum()
        er = eng / views if views > 0 else None
    else:
        eng = x['engagement'].sum()
        foll = x['followers'].sum()
        er = eng / foll if foll > 0 else None

    return pd.Series({
        'Post Count': len(x),
        'Avg ER%': er
    })

# -----------------------------------------
# CONTENT SCOPES
# -----------------------------------------

scopes = {
    'Paid': df[df['is_paid']],
    'Organic': df[df['post_type'] == 'Organic'],
    'Branded': df[df['post_type'] == 'Branded'],
    'Influencer': df[df['post_type'] == 'Influencer'],
    'Creator': df[df['post_type'] == 'Creator']
}

outputs = {}

# -----------------------------------------
# GENERATE TABLES
# -----------------------------------------

for scope, sdf in scopes.items():

    overall = sdf.groupby('brand').apply(overall_metrics).reset_index()
    outputs[f"{scope}_Brand_Overall"] = sort_brands(overall, 'brand', 'Total Posts')

    fmt = sdf.groupby(['brand', 'format']).apply(format_metrics).reset_index()
    outputs[f"{scope}_Brand_Format"] = sort_brands(fmt, 'brand', 'Post Count')

    typ = sdf.groupby(['brand', 'post_type']).apply(overall_metrics).reset_index()
    outputs[f"{scope}_Brand_Type"] = sort_brands(typ, 'brand', 'Total Posts')

    src = sdf.groupby(['brand', 'platform']).apply(overall_metrics).reset_index()
    outputs[f"{scope}_Brand_Source"] = sort_brands(src, 'brand', 'Total Posts')

    wk = sdf.groupby(['brand', 'Week']).apply(overall_metrics).reset_index()
    outputs[f"{scope}_Brand_Weekly"] = wk

# -----------------------------------------
# INFLUENCER TIER ANALYSIS
# -----------------------------------------

influencer_df = df[df['post_type'] == 'Influencer']

tier_table = (
    influencer_df
    .groupby(['brand', 'influencer_tier'])
    .apply(overall_metrics)
    .reset_index()
)

outputs['Influencer_Tier_Analysis'] = sort_brands(
    tier_table, 'brand', 'Total Posts'
)

# -----------------------------------------
# EXPORT EXCEL
# -----------------------------------------

with pd.ExcelWriter(EXCEL_OUTPUT, engine='openpyxl') as writer:
    for name, table in outputs.items():
        table.to_excel(writer, sheet_name=name[:31], index=False)

# -----------------------------------------
# AUTO-GENERATE PPT – WoW TRENDLINES
# -----------------------------------------

prs = Presentation()

def add_trend_slide(df, title):
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = title

    chart_data = CategoryChartData()
    weeks = sorted(df['Week'].dropna().unique())
    chart_data.categories = weeks

    for brand in PRIORITY_BRANDS:
        bdf = df[df['brand'] == brand].sort_values('Week')
        if not bdf.empty:
            chart_data.add_series(
                brand,
                bdf['Avg ER%'].fillna(0).tolist()
            )

    slide.shapes.add_chart(
        XL_CHART_TYPE.LINE,
        x=0,
        y=0,
        cx=9144000,
        cy=4572000,
        chart_data=chart_data
    )

add_trend_slide(outputs['Paid_Brand_Weekly'], 'Paid Content – WoW Avg ER%')
add_trend_slide(outputs['Branded_Brand_Weekly'], 'Branded Content – WoW Avg ER%')
add_trend_slide(outputs['Influencer_Brand_Weekly'], 'Influencer Content – WoW Avg ER%')
add_trend_slide(outputs['Creator_Brand_Weekly'], 'Creator Content – WoW Avg ER%')

prs.save(PPT_OUTPUT)

print("PROCESS COMPLETED SUCCESSFULLY")
print("Excel Output:", EXCEL_OUTPUT)
print("PPT Output:", PPT_OUTPUT)

