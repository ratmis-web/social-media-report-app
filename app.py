# =========================================================
# SOCIAL MEDIA REPORTING AUTOMATION – FINAL MULTI-FILE OUTPUT
# =========================================================

import streamlit as st
import pandas as pd
import numpy as np

# ---------------------------------------------------------
# STREAMLIT SETUP
# ---------------------------------------------------------

st.set_page_config(page_title="Social Media Reporting", layout="wide")
st.title("📊 Social Media Reporting Automation")

uploaded_file = st.file_uploader("Upload input Excel (.xlsx)", type=["xlsx"])
if uploaded_file is None:
    st.stop()

# ---------------------------------------------------------
# CONFIG
# ---------------------------------------------------------

PRIORITY_BRANDS = [
    "Similac", "Ensure", "Pediasure",
    "Pedialyte", "Juven", "Glucerna"
]

# ---------------------------------------------------------
# LOAD + CLEAN DATA
# ---------------------------------------------------------

df = pd.read_excel(uploaded_file)

df = df.rename(columns={
    "Brand Name/Category Name": "brand",
    "Social Channel/Platform": "platform",
    "Type of Post (Branded, Influencer,Creators, Organic, Shop)": "post_type",
    "Post Format (Video, Reels, Shorts, Images, Carousels)": "format",
    "Video Plays": "views",
    "Followers": "followers",
    "Influencer Tier": "influencer_tier"
})

df["engagement"] = (
    pd.to_numeric(df["Likes"], errors="coerce").fillna(0) +
    pd.to_numeric(df["Comments"], errors="coerce").fillna(0)
)

df["views"] = pd.to_numeric(df["views"], errors="coerce").fillna(0)
df["followers"] = pd.to_numeric(df["followers"], errors="coerce").fillna(0)

df["brand"] = df["brand"].astype(str).str.strip().str.title()

df["post_type"] = (
    df["post_type"]
    .astype(str)
    .str.title()
    .replace({"Tagged": "Influencer", "Creators": "Creator"})
)

df["format"] = (
    df["format"]
    .astype(str)
    .str.title()
    .replace({
        "Images": "Image",
        "Reels": "Shorts",
        "Sidecar": "Carousel",
        "Image+Video": "Carousel"
    })
)

df["is_dynamic"] = df["format"].isin(["Video", "Shorts"])
df["is_paid"] = df["post_type"].isin(["Branded", "Influencer", "Creator", "Shop"])

# ---------------------------------------------------------
# HELPERS
# ---------------------------------------------------------

def brand_sort(df, volume_col):
    df = df.copy()
    df["_priority"] = df["brand"].apply(
        lambda x: PRIORITY_BRANDS.index(x)
        if x in PRIORITY_BRANDS else len(PRIORITY_BRANDS)
    )
    df = df.sort_values(by=["_priority", volume_col], ascending=[True, False])
    return df.drop(columns="_priority")

def calculate_er(x):
    if x["is_dynamic"].any():
        v = x.loc[x["is_dynamic"], "views"].sum()
        e = x.loc[x["is_dynamic"], "engagement"].sum()
        return e / v if v > 0 else None
    else:
        f = x["followers"].sum()
        e = x["engagement"].sum()
        return e / f if f > 0 else None

def generate_tables(data, dimension):
    # Post Count
    post_count = (
        data.groupby(["brand", dimension])
        .size()
        .reset_index(name="Post Count")
        .pivot(index="brand", columns=dimension, values="Post Count")
        .fillna(0)
        .reset_index()
    )

    # Avg ER
    avg_er = (
        data.groupby(["brand", dimension])
        .apply(calculate_er)
        .reset_index(name="Avg ER%")
        .pivot(index="brand", columns=dimension, values="Avg ER%")
        .reset_index()
        .fillna("-")
    )

    post_count = brand_sort(post_count, post_count.columns[1])
    avg_er = brand_sort(avg_er, avg_er.columns[1])

    return post_count, avg_er

# ---------------------------------------------------------
# CONTENT SCOPES
# ---------------------------------------------------------

content_scopes = {
    "Paid": df[df["is_paid"]],
    "Organic": df[df["post_type"] == "Organic"],
    "Branded": df[df["post_type"] == "Branded"],
    "Influencer": df[df["post_type"] == "Influencer"],
    "Creator": df[df["post_type"] == "Creator"]
}

# ---------------------------------------------------------
# EXPORT EACH SCOPE AS SEPARATE FILE
# ---------------------------------------------------------

download_files = {}

for scope, sdf in content_scopes.items():
    file_name = f"{scope}_Content_Report.xlsx"

    with pd.ExcelWriter(file_name, engine="openpyxl") as writer:
        for dim, label in [("format", "Format"), ("post_type", "Type"), ("platform", "Source")]:
            pc, er = generate_tables(sdf, dim)
            pc.to_excel(writer, sheet_name=f"{scope}_Brand_{label}_PostCount", index=False)
            er.to_excel(writer, sheet_name=f"{scope}_Brand_{label}_AvgER", index=False)

    download_files[scope] = file_name

# ---------------------------------------------------------
# INFLUENCER TIER ANALYSIS
# ---------------------------------------------------------

tier_df = df[df["post_type"] == "Influencer"]

tier_post_count = (
    tier_df.groupby(["brand", "influencer_tier"])
    .size()
    .reset_index(name="Post Count")
    .pivot(index="brand", columns="influencer_tier", values="Post Count")
    .fillna(0)
    .reset_index()
)

tier_avg_er = (
    tier_df.groupby(["brand", "influencer_tier"])
    .apply(calculate_er)
    .reset_index(name="Avg ER%")
    .pivot(index="brand", columns="influencer_tier", values="Avg ER%")
    .fillna("-")
    .reset_index()
)

tier_post_count = brand_sort(tier_post_count, tier_post_count.columns[1])
tier_avg_er = brand_sort(tier_avg_er, tier_avg_er.columns[1])

tier_file = "Influencer_Tier_Report.xlsx"
with pd.ExcelWriter(tier_file, engine="openpyxl") as writer:
    tier_post_count.to_excel(writer, sheet_name="Tier_PostCount", index=False)
    tier_avg_er.to_excel(writer, sheet_name="Tier_AvgER", index=False)

download_files["Influencer_Tier"] = tier_file

# ---------------------------------------------------------
# DOWNLOAD SECTION
# ---------------------------------------------------------

st.success("All reports generated successfully 🎉")

for label, file in download_files.items():
    with open(file, "rb") as f:
        st.download_button(
            f"📥 Download {label} Report",
            f,
            file_name=file
        )