import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(page_title="Social Media Reporting Automation", layout="wide")

st.title("📊 Social Media Performance Automation")
st.markdown(
    "Upload your **raw social media .xlsx file** and download a **ready-to-use Excel report**."
)

# ------------------------
# File Upload
# ------------------------
uploaded_file = st.file_uploader(
    "Upload raw data file (.xlsx)", type=["xlsx"]
)

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # ------------------------
    # Data Cleaning & Standardization
    # ------------------------
    df = df.rename(columns={
        'Social Channel/Platform': 'platform',
        'URL': 'post_url',
        'Brand Name/Category Name': 'brand',
        'Type of Post (Branded, Influencer,Creators, Organic, Shop)': 'post_type',
        'Post Format (Video, Reels, Shorts, Images, Carousels)': 'format',
        'Followers': 'followers',
        'Likes': 'likes',
        'Comments': 'comments',
        'Video Plays': 'views'
    })

    for col in ['likes', 'comments', 'views', 'followers']:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    df['engagement'] = df['likes'] + df['comments']

    df['brand'] = (
        df['brand'].astype(str)
        .str.strip()
        .str.title()
        .replace({'Nan': 'NAN', 'Quaker Oats': 'Quaker'})
    )

    df['post_type'] = (
        df['post_type'].str.strip()
        .str.title()
        .replace({'Tagged': 'Influencer', 'Influencers': 'Influencer', 'Creators': 'Creator'})
    )

    df['format'] = (
        df['format'].str.strip()
        .str.title()
        .replace({
            'Sidecar': 'Carousel',
            'Image+Video': 'Carousel',
            'Images': 'Image',
            'Videos': 'Video',
            'Reels': 'Shorts'
        })
    )

    df['is_video'] = df['format'].isin(['Video', 'Shorts'])

    # ------------------------
    # Brand Overall Table
    # ------------------------
    brand_overall = (
        df.groupby('brand', as_index=False)
        .agg(
            total_posts=('post_url', 'count'),
            total_engagement=('engagement', 'sum'),
            video_engagement=('engagement', lambda x: x[df.loc[x.index, 'is_video']].sum()),
            video_views=('views', lambda x: x[df.loc[x.index, 'is_video']].sum())
        )
    )

    brand_overall['Avg ER'] = (
        brand_overall['video_engagement'] /
        brand_overall['video_views']
    ).replace([np.inf, np.nan], 'No video posts')

    brand_overall['Avg ER'] = brand_overall['Avg ER'].apply(
        lambda x: f"{x:.2%}" if isinstance(x, float) else x
    )

    brand_overall = brand_overall.rename(columns={
        'brand': 'Brand',
        'total_posts': 'Total Posts',
        'total_engagement': 'Total Engagement',
        'video_views': 'Total Views'
    })

    # ------------------------
    # Helper: Format-wise ER
    # ------------------------
    df['post_er'] = np.where(
        df['format'].isin(['Video', 'Shorts']),
        df['engagement'] / df['views'],
        np.where(
            df['format'].isin(['Image', 'Carousel']),
            df['engagement'] / df['followers'],
            np.nan
        )
    )

    # ------------------------
    # Create Excel Output
    # ------------------------
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        brand_overall.to_excel(writer, sheet_name='Brand_Overall', index=False)

        df.groupby(['brand', 'format']).size().unstack(fill_value=0).reset_index() \
          .to_excel(writer, 'Format_Post_Count', index=False)

        df.groupby(['brand', 'post_type']).size().unstack(fill_value=0).reset_index() \
          .to_excel(writer, 'PostType_Post_Count', index=False)

        df.groupby(['brand', 'platform']).size().unstack(fill_value=0).reset_index() \
          .to_excel(writer, 'Platform_Post_Count', index=False)

        df.groupby(['brand', 'format'])['post_er'].mean().unstack().reset_index() \
          .to_excel(writer, 'Format_Avg_ER', index=False)

        df.groupby(['brand', 'post_type'])['post_er'].mean().unstack().reset_index() \
          .to_excel(writer, 'PostType_Avg_ER', index=False)

        df.groupby(['brand', 'platform'])['post_er'].mean().unstack().reset_index() \
          .to_excel(writer, 'Platform_Avg_ER', index=False)

    st.success("Report generated successfully!")

    st.download_button(
        label="📥 Download Excel Report",
        data=output.getvalue(),
        file_name="Final_Social_Media_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )