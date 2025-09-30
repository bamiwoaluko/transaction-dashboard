import streamlit as st
import pandas as pd
import plotly.express as px
import os
import io
import zipfile


def df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    """Convert DataFrame to Excel bytes for writing into zip."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False)
    return buffer.getvalue()


st.set_page_config(page_title="Transaction Comparison", page_icon=":bar_chart:", layout="wide")
st.title(":bar_chart: Transaction Comparison")
st.markdown('<style>div.block-container{padding-top:2rem;}</style>', unsafe_allow_html=True)

# Upload files
st.sidebar.header("Upload Files")
week1_file = st.sidebar.file_uploader("Upload Previous Week file", type=["xlsx"], key="Prev Week")
week2_file = st.sidebar.file_uploader("Upload Current Week file", type=["xlsx"], key="Curr Week")

#Utitlity functions
def load_data(file, sheet_name):
    """Load Excel sheet into DataFrame and drop the TOTAL row if present."""
    df = pd.read_excel(file, sheet_name=sheet_name)
    if str(df.iloc[-1, 0]).strip() == "TOTAL":
        df = df.iloc[:-1]
    return df

def get_week_label(uploaded_file, fallback):
    """Extract the first part of filename (e.g., WEEK_4) else use fallback."""
    if not uploaded_file:
        return fallback
    base = os.path.basename(uploaded_file.name)
    first_part = base.split()[0]  
    return first_part.upper() if first_part else fallback

# Get dynamic labels
week1_label = get_week_label(week1_file, "PREVIOUS_WEEK")
week2_label = get_week_label(week2_file, "CURRENT_WEEK")

# weekly totals comparison
def compare_totals(week1_file, week2_file):
    results = []
    for sheet in ["WITH 6010", "WITHOUT 6010"]:
        df1 = load_data(week1_file, sheet)
        df2 = load_data(week2_file, sheet)

        week1_volume = round(df1["VOLUME"].sum())
        week1_value = round(df1["VALUE"].sum())
        week2_volume = round(df2["VOLUME"].sum())
        week2_value = round(df2["VALUE"].sum())

        vol_diff = week2_volume - week1_volume
        val_diff = week2_value - week1_value

        vol_pct = round((vol_diff / week1_volume * 100), 2) if week1_volume else 0
        val_pct = round((val_diff / week1_value * 100), 2) if week1_value else 0

        results.append({
            "SHEET": sheet,
            f"{week1_label} VOLUME": f"{week1_volume:,}",
            f"{week2_label} VOLUME": f"{week2_volume:,}",
            "VOLUME DIFF": f"{vol_diff:,}",
            "VOLUME %": f"{vol_pct}%",
            f"{week1_label} VALUE": f"{week1_value:,}",
            f"{week2_label} VALUE": f"{week2_value:,}",
            "VALUE DIFF": f"{val_diff:,}",
            "VALUE %": f"{val_pct}%",
        })

    return pd.DataFrame(results)

if week1_file and week2_file:
    comparison = compare_totals(week1_file, week2_file)

    st.subheader("📊 Weekly Totals Comparison")
    st.dataframe(
        comparison.style.format(na_rep="-")
    )

    #bar chart side by side
    st.subheader("🔍 Visual Comparison")
    melted = comparison.melt(id_vars=["SHEET"])
    melted.columns = ["SHEET", "Metric", "Amount"]
    melted["Amount"] = melted["Amount"].replace({",": ""}, regex=True).str.replace("%", "", regex=False)
    melted["Amount"] = pd.to_numeric(melted["Amount"], errors="coerce")

    volume_metrics = [f"{week1_label} VOLUME", f"{week2_label} VOLUME"]
    value_metrics = [f"{week1_label} VALUE", f"{week2_label} VALUE"]

    # Volume chart
    fig_vol = px.bar(
        melted[melted["Metric"].isin(volume_metrics)],
        x="SHEET", y="Amount", color="Metric", barmode="group",
        title="Volume Comparison", text="Amount"
    )
    fig_vol.update_traces(texttemplate="%{text:,.0f}", textposition="outside")

    # Value chart
    fig_val = px.bar(
        melted[melted["Metric"].isin(value_metrics)],
        x="SHEET", y="Amount", color="Metric", barmode="group",
        title="Value Comparison", text="Amount"
    )
    fig_val.update_traces(texttemplate="%{text:,.0f}", textposition="outside")

    col1, col2 = st.columns(2)
    with col1: st.plotly_chart(fig_vol, use_container_width=True)
    with col2: st.plotly_chart(fig_val, use_container_width=True)


#Region and Bank Weekly comparison

def region_bank_comparison(week1_file, week2_file):
    df1 = load_data(week1_file, "WITH 6010")
    df2 = load_data(week2_file, "WITH 6010")

    agg1 = df1.groupby(["REGION", "BANK"], as_index=False)[["VOLUME", "VALUE"]].sum()
    agg2 = df2.groupby(["REGION", "BANK"], as_index=False)[["VOLUME", "VALUE"]].sum()

    merged = pd.merge(agg1, agg2, on=["REGION", "BANK"], suffixes=(f"_{week1_label}", f"_{week2_label}"), how="outer").fillna(0)

    merged["VOLUME_DIFF"] = merged[f"VOLUME_{week2_label}"] - merged[f"VOLUME_{week1_label}"]
    merged["VALUE_DIFF"] = merged[f"VALUE_{week2_label}"] - merged[f"VALUE_{week1_label}"]

    merged["VOLUME_%"] = merged.apply(
        lambda x: round((x["VOLUME_DIFF"] / x[f"VOLUME_{week1_label}"] * 100), 2) if x[f"VOLUME_{week1_label}"] else 0, axis=1
    )
    merged["VALUE_%"] = merged.apply(
        lambda x: round((x["VALUE_DIFF"] / x[f"VALUE_{week1_label}"] * 100), 2) if x[f"VALUE_{week1_label}"] else 0, axis=1
    )

    for col in [f"VOLUME_{week1_label}", f"VOLUME_{week2_label}", "VOLUME_DIFF",
                f"VALUE_{week1_label}", f"VALUE_{week2_label}", "VALUE_DIFF"]:
        merged[col] = merged[col].round(0).astype(int)

    return merged

if week1_file and week2_file:
    comparison = region_bank_comparison(week1_file, week2_file)

    st.subheader("📊 Region + Bank Weekly Comparison")
    st.dataframe(
        comparison.style.format(na_rep="-")
    )

# Region analysis

if week1_file and week2_file:
    df1 = load_data(week1_file, "WITH 6010")
    df2 = load_data(week2_file, "WITH 6010")

    df1["REGION"] = df1["REGION"].fillna("UNKNOWN")
    df2["REGION"] = df2["REGION"].fillna("UNKNOWN")

    week1_grouped = df1.groupby("REGION")[["VOLUME", "VALUE"]].sum().reset_index()
    week2_grouped = df2.groupby("REGION")[["VOLUME", "VALUE"]].sum().reset_index()

    region_comparison = pd.merge(
        week1_grouped, week2_grouped,
        on="REGION", how="outer",
        suffixes=(f"_{week1_label}", f"_{week2_label}")
    ).fillna(0)

    region_comparison["VOL_DIFF"] = region_comparison[f"VOLUME_{week2_label}"] - region_comparison[f"VOLUME_{week1_label}"]
    region_comparison["VAL_DIFF"] = region_comparison[f"VALUE_{week2_label}"] - region_comparison[f"VALUE_{week1_label}"]
    region_comparison["VOL_%CHANGE"] = ((region_comparison["VOL_DIFF"] / region_comparison[f"VOLUME_{week1_label}"]) * 100).replace([float("inf"), -float("inf")], 0)
    region_comparison["VAL_%CHANGE"] = ((region_comparison["VAL_DIFF"] / region_comparison[f"VALUE_{week1_label}"]) * 100).replace([float("inf"), -float("inf")], 0)

    # just volume table
    volume_table = region_comparison[["REGION", f"VOLUME_{week1_label}", f"VOLUME_{week2_label}"]].copy()
    volume_total = pd.DataFrame([{
        "REGION": "TOTAL",
        f"VOLUME_{week1_label}": volume_table[f"VOLUME_{week1_label}"].sum(),
        f"VOLUME_{week2_label}": volume_table[f"VOLUME_{week2_label}"].sum()
    }])
    volume_table = pd.concat([volume_table, volume_total], ignore_index=True)

    st.subheader("📊 Region Volume Comparison")
    st.dataframe(volume_table.style.format({
        f"VOLUME_{week1_label}": "{:,.0f}", f"VOLUME_{week2_label}": "{:,.0f}"
    }))

    # just value table
    value_table = region_comparison[["REGION", f"VALUE_{week1_label}", f"VALUE_{week2_label}"]].copy()
    value_total = pd.DataFrame([{
        "REGION": "TOTAL",
        f"VALUE_{week1_label}": value_table[f"VALUE_{week1_label}"].sum(),
        f"VALUE_{week2_label}": value_table[f"VALUE_{week2_label}"].sum()
    }])
    value_table = pd.concat([value_table, value_total], ignore_index=True)

    st.subheader("💰 Region Value Comparison")
    st.dataframe(value_table.style.format({
        f"VALUE_{week1_label}": "{:,.0f}", f"VALUE_{week2_label}": "{:,.0f}"
    }))

    #difference table
    diff_table = region_comparison[["REGION", "VOL_DIFF", "VOL_%CHANGE", "VAL_DIFF", "VAL_%CHANGE"]].copy()
    diff_total = pd.DataFrame([{
        "REGION": "TOTAL",
        "VOL_DIFF": diff_table["VOL_DIFF"].sum(),
        "VOL_%CHANGE": diff_table["VOL_%CHANGE"].sum(),
        "VAL_DIFF": diff_table["VAL_DIFF"].sum(),
        "VAL_%CHANGE": diff_table["VAL_%CHANGE"].sum()
    }])
    diff_table = pd.concat([diff_table, diff_total], ignore_index=True)

    st.subheader("🔄 Region Difference")
    st.dataframe(diff_table.style.format({
        "VOL_DIFF": "{:,.0f}", "VOL_%CHANGE": "{:.2f}%",
        "VAL_DIFF": "{:,.0f}", "VAL_%CHANGE": "{:.2f}%"
    }))


    fig_vol = px.bar(
        volume_table[volume_table["REGION"] != "TOTAL"],
        x="REGION",
        y=[f"VOLUME_{week1_label}", f"VOLUME_{week2_label}"],
        barmode="group",
        title="Region Volume Comparison (WITH 6010)"
    )
    fig_vol.update_traces(texttemplate="%{y:,.0f}", textposition="outside")
    st.plotly_chart(fig_vol, use_container_width=True)

    fig_val = px.bar(
        value_table[value_table["REGION"] != "TOTAL"],
        x="REGION",
        y=[f"VALUE_{week1_label}", f"VALUE_{week2_label}"],
        barmode="group",
        title="Region Value Comparison (WITH 6010)"
    )
    fig_val.update_traces(texttemplate="%{y:,.0f}", textposition="outside")
    st.plotly_chart(fig_val, use_container_width=True)

#State level analysis

state_mapping = {
    "AB": "Abia", "AD": "Adamawa", "AK": "Akwa Ibom", "AN": "Anambra", "BA": "Bauchi",
    "BY": "Bayelsa", "BE": "Benue", "BO": "Borno", "CR": "Cross River", "DE": "Delta",
    "EB": "Ebonyi", "ED": "Edo", "EK": "Ekiti", "EN": "Enugu", "FC": "FCT - Abuja",
    "GO": "Gombe", "IM": "Imo", "JI": "Jigawa", "KD": "Kaduna", "KN": "Kano",
    "KT": "Katsina", "KE": "Kebbi", "KO": "Kogi", "KW": "Kwara", "LA": "Lagos",
    "NA": "Nasarawa", "NI": "Niger", "OG": "Ogun", "ON": "Ondo", "OS": "Osun",
    "OY": "Oyo", "PL": "Plateau", "RI": "Rivers", "SO": "Sokoto", "TA": "Taraba",
    "YO": "Yobe", "ZA": "Zamfara"
}

if week1_file and week2_file:
    df1 = load_data(week1_file, "WITH 6010")
    df2 = load_data(week2_file, "WITH 6010")

    df1["STATE"] = df1["STATE"].fillna("UNKNOWN")
    df2["STATE"] = df2["STATE"].fillna("UNKNOWN")

    df1["STATE"] = df1["STATE"].map(state_mapping).fillna(df1["STATE"])
    df2["STATE"] = df2["STATE"].map(state_mapping).fillna(df2["STATE"])

    week1_grouped = df1.groupby("STATE")[["VOLUME", "VALUE"]].sum().reset_index()
    week2_grouped = df2.groupby("STATE")[["VOLUME", "VALUE"]].sum().reset_index()

    state_comparison = pd.merge(
        week1_grouped, week2_grouped,
        on="STATE", how="outer",
        suffixes=(f"_{week1_label}", f"_{week2_label}")
    ).fillna(0)

    state_comparison["VOL_DIFF"] = state_comparison[f"VOLUME_{week2_label}"] - state_comparison[f"VOLUME_{week1_label}"]
    state_comparison["VAL_DIFF"] = state_comparison[f"VALUE_{week2_label}"] - state_comparison[f"VALUE_{week1_label}"]
    state_comparison["VOL_%CHANGE"] = ((state_comparison["VOL_DIFF"] / state_comparison[f"VOLUME_{week1_label}"]) * 100).replace([float("inf"), -float("inf")], 0)
    state_comparison["VAL_%CHANGE"] = ((state_comparison["VAL_DIFF"] / state_comparison[f"VALUE_{week1_label}"]) * 100).replace([float("inf"), -float("inf")], 0)

    state_comparison = state_comparison[["STATE", "VOL_DIFF", "VOL_%CHANGE", "VAL_DIFF", "VAL_%CHANGE"]]

    state_comparison.loc[len(state_comparison)] = {
        "STATE": "TOTAL",
        "VOL_DIFF": state_comparison["VOL_DIFF"].sum(),
        "VOL_%CHANGE": state_comparison["VOL_%CHANGE"].sum(),
        "VAL_DIFF": state_comparison["VAL_DIFF"].sum(),
        "VAL_%CHANGE": state_comparison["VAL_%CHANGE"].sum()
    }

    st.subheader("📑 State Comparison")
    st.dataframe(state_comparison.style.format({
        "VOL_DIFF": "{:,.0f}", "VOL_%CHANGE": "{:.2f}%",
        "VAL_DIFF": "{:,.0f}", "VAL_%CHANGE": "{:.2f}%"
    }))

    #download tables as zip file
    if week1_file and week2_file:
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zf:
            # Add all key tables
            zf.writestr("weekly_totals.xlsx", df_to_excel_bytes(compare_totals(week1_file, week2_file)))
            zf.writestr("region_bank_comparison.xlsx", df_to_excel_bytes(region_bank_comparison(week1_file, week2_file)))
            zf.writestr("region_volume.xlsx", df_to_excel_bytes(volume_table))
            zf.writestr("region_value.xlsx", df_to_excel_bytes(value_table))
            zf.writestr("region_differences.xlsx", df_to_excel_bytes(diff_table))
            zf.writestr("state_comparison.xlsx", df_to_excel_bytes(state_comparison))

        st.download_button(
            "⬇️ Download All Tables (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="all_tables.zip",
            mime="application/zip"
        )
