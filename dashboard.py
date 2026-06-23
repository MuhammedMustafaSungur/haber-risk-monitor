import os
import re
from PIL import Image
import streamlit.components.v1 as components
import pandas as pd
import streamlit as st
import plotly.express as px

DATA_PATH = os.path.join("data", "haberler_final.csv")
LOGO_PATH = os.path.join("assets", "logo.png")

page_icon = "📰"
if os.path.exists(LOGO_PATH):
    page_icon = Image.open(LOGO_PATH)

st.set_page_config(
    page_title="Haber Risk Monitor",
    page_icon=page_icon,
    layout="wide"
)

st.markdown("""
<style>
.stApp {
    background:
        linear-gradient(rgba(2, 6, 23, 0.91), rgba(15, 23, 42, 0.96)),
        url("https://images.unsplash.com/photo-1642790106117-e829e14a795f?auto=format&fit=crop&w=2200&q=80");
    background-size: cover;
    background-position: center;
    background-attachment: fixed;
    color: #e5e7eb;
}

.stApp::before {
    content: "";
    position: fixed;
    inset: 0;
    background:
        radial-gradient(circle at 15% 20%, rgba(56, 189, 248, 0.14), transparent 28%),
        radial-gradient(circle at 80% 10%, rgba(34, 197, 94, 0.10), transparent 30%),
        radial-gradient(circle at 50% 85%, rgba(99, 102, 241, 0.16), transparent 35%);
    backdrop-filter: blur(3px);
    z-index: -1;
}

[data-testid="stSidebar"] {
    background: rgba(15, 23, 42, 0.96);
}

.block-container {
    padding-top: 2rem;
}

.kpi-card {
    background: linear-gradient(135deg, rgba(30,41,59,.92), rgba(15,23,42,.92));
    border: 1px solid rgba(148, 163, 184, 0.22);
    border-radius: 18px;
    padding: 18px;
    min-height: 112px;
    box-shadow: 0 20px 45px rgba(0,0,0,0.22);
}

.kpi-title {
    color: #94a3b8;
    font-size: 13px;
}

.kpi-value {
    color: #f8fafc;
    font-size: 32px;
    font-weight: 800;
}

.news-card {
    background: rgba(15,23,42,0.80);
    border: 1px solid rgba(148, 163, 184, 0.22);
    border-radius: 18px;
    padding: 18px;
    margin-bottom: 14px;
}

.badge {
    display: inline-block;
    padding: 4px 10px;
    margin: 3px;
    border-radius: 999px;
    background: rgba(56,189,248,0.13);
    border: 1px solid rgba(56,189,248,0.25);
    color: #bae6fd;
    font-size: 12px;
}

a {
    color: #38bdf8 !important;
}

table {
    width: 100%;
    border-collapse: collapse;
    font-size: 13px;
}

th {
    background: rgba(30, 41, 59, 0.95);
    color: #e2e8f0;
    padding: 8px;
}

td {
    background: rgba(15, 23, 42, 0.72);
    color: #dbeafe;
    border-bottom: 1px solid rgba(148, 163, 184, 0.12);
    padding: 8px;
}
</style>
""", unsafe_allow_html=True)


def clean_text(x):
    return re.sub(r"\s+", " ", str(x)).strip()


def ensure_col(df, col):
    if col not in df.columns:
        df[col] = ""
    return df


def split_multi_value(value):
    value = clean_text(value)
    if not value:
        return []
    parts = re.split(r"\s*\|\s*|,\s*|;\s*", value)
    return [clean_text(p) for p in parts if clean_text(p)]


def unique_split_values(series):
    values = []
    for x in series:
        values.extend(split_multi_value(x))
    return sorted(set(values))


def filter_multi_col(dataframe, col, selected):
    if not selected:
        return dataframe

    selected_set = set(selected)

    return dataframe[
        dataframe[col].apply(
            lambda x: bool(selected_set.intersection(split_multi_value(x)))
        )
    ]


def explode_count(dataframe, col, label, top_n=15):
    rows = []
    for value in dataframe[col]:
        for item in split_multi_value(value):
            rows.append(item)

    if not rows:
        return pd.DataFrame(columns=[label, "Haber Sayısı"])

    out = (
        pd.Series(rows)
        .value_counts()
        .head(top_n)
        .reset_index()
    )
    out.columns = [label, "Haber Sayısı"]
    return out


def company_display(row):
    company = clean_text(row.get("company_candidates", ""))
    maps = clean_text(row.get("maps_best_guess", ""))
    return company if company else maps


def short_text(text, max_len=260):
    text = clean_text(text)
    return text if len(text) <= max_len else text[:max_len].rstrip() + "..."


HEADER_LOGO = os.path.join("assets", "vakif_katilim.png")

top_left, center, top_right = st.columns([5, 2, 2])

with top_left:
    st.image(HEADER_LOGO, width=500)

with top_right:
    st.image(LOGO_PATH, width=220)

st.markdown("""
<h1 style="
text-align:center;
font-size:72px;
font-weight:900;
color:white;
margin-top:0px;
margin-bottom:-5px;
padding:0;
line-height:1;">
MERCEK
</h1>
""", unsafe_allow_html=True)

st.markdown("""
<div style="
    width:420px;
    height:4px;
    background:linear-gradient(
        90deg,
        rgba(230,0,126,0),
        #e6007e,
        rgba(230,0,126,0)
    );
    margin:0px auto 4px auto;
">
</div>
""", unsafe_allow_html=True)

st.markdown("""
<p style="
text-align:center;
font-size:24px;
color:white;
margin-top:0px;
margin-bottom:0px;
padding:0;
line-height:1.1;">
Medya Erken Risk Kontrol ve İzleme Platformu
</p>
""", unsafe_allow_html=True)

st.markdown("""
<p style="
text-align:center;
font-size:20px;
font-weight:700;
color:#ff4da6;
margin-top:4px;
margin-bottom:0px;
padding:0;
line-height:1.1;">
Kredi İzleme Müdürlüğü
</p>
""", unsafe_allow_html=True)


st.divider()

if not os.path.exists(DATA_PATH):
    st.error("data/haberler_final.csv bulunamadı.")
    st.stop()

df = pd.read_csv(DATA_PATH).fillna("")

if df.empty:
    st.warning("CSV boş.")
    st.stop()

required_cols = [
    "published", "category", "title", "summary", "content", "raw_text",
    "city_candidates", "district_candidates", "neighborhood_candidates",
    "company_candidates", "activity_candidates", "maps_best_guess",
    "maps_search_query", "link"
]

for col in required_cols:
    df = ensure_col(df, col)

df["published_dt"] = pd.to_datetime(df["published"], errors="coerce")
df["firma_display"] = df.apply(company_display, axis=1)

st.sidebar.header("Filtreler")

categories = sorted([x for x in df["category"].unique() if clean_text(x)])
cities = unique_split_values(df["city_candidates"])
districts = unique_split_values(df["district_candidates"])
firms = sorted([x for x in df["firma_display"].unique() if clean_text(x)])

selected_category = st.sidebar.multiselect("Kategori", categories)
selected_city = st.sidebar.multiselect("İl", cities)
selected_district = st.sidebar.multiselect("İlçe", districts)
selected_firm = st.sidebar.multiselect("Firma / Maps Adayı", firms)
search_text = st.sidebar.text_input("Başlık / firma / il / içerik ara")

filtered = df.copy()

if selected_category:
    filtered = filtered[filtered["category"].isin(selected_category)]

filtered = filter_multi_col(filtered, "city_candidates", selected_city)
filtered = filter_multi_col(filtered, "district_candidates", selected_district)

if selected_firm:
    filtered = filtered[filtered["firma_display"].isin(selected_firm)]

if search_text:
    s = search_text.lower().strip()
    filtered = filtered[
        filtered["title"].str.lower().str.contains(s, na=False)
        | filtered["summary"].str.lower().str.contains(s, na=False)
        | filtered["content"].str.lower().str.contains(s, na=False)
        | filtered["raw_text"].str.lower().str.contains(s, na=False)
        | filtered["firma_display"].str.lower().str.contains(s, na=False)
        | filtered["city_candidates"].str.lower().str.contains(s, na=False)
        | filtered["district_candidates"].str.lower().str.contains(s, na=False)
    ]

filtered = filtered.sort_values("published_dt", ascending=False, na_position="last")

k1, k2, k3, k4, k5 = st.columns(5)

kpis = [
    ("Toplam Haber", len(df)),
    ("Filtreli Haber", len(filtered)),
    ("Kategori", filtered["category"].replace("", pd.NA).dropna().nunique()),
    ("İl", len(unique_split_values(filtered["city_candidates"]))),
    ("Firma / Maps", filtered["firma_display"].replace("", pd.NA).dropna().nunique()),
]

for col, (title, value) in zip([k1, k2, k3, k4, k5], kpis):
    with col:
        st.markdown(
            f"""
            <div class="kpi-card">
                <div class="kpi-title">{title}</div>
                <div class="kpi-value">{value}</div>
            </div>
            """,
            unsafe_allow_html=True
        )

st.divider()

c1, c2 = st.columns(2)

with c1:
    cat_count = filtered[filtered["category"] != ""]["category"].value_counts().reset_index()
    cat_count.columns = ["Kategori", "Haber Sayısı"]

    if not cat_count.empty:
        total = cat_count["Haber Sayısı"].sum()
        cat_count["Yüzde"] = cat_count["Haber Sayısı"] / total * 100
        cat_count["Etiket"] = cat_count.apply(
            lambda r: f"{r['Kategori']}<br>%{r['Yüzde']:.1f}",
            axis=1
        )

        fig = px.treemap(
            cat_count,
            path=["Kategori"],
            values="Haber Sayısı",
            title="Risk Kategorileri",
            custom_data=["Yüzde", "Haber Sayısı"]
        )

        fig.update_traces(
            texttemplate="<b>%{label}</b><br>%{percentRoot:.1%}<br>%{value} haber",
            textfont_size=22,
            textinfo="label+text+value",
            marker=dict(line=dict(width=2, color="#0f172a"))
        )

        fig.update_layout(
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            font_color="#e5e7eb",
            title_font_size=22,
            margin=dict(t=50, l=10, r=10, b=10)
        )

        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Kategori verisi yok.")

with c2:
    city_count = explode_count(filtered, "city_candidates", "İl", top_n=12)

    if not city_count.empty:
        fig = px.bar(
            city_count,
            x="Haber Sayısı",
            y="İl",
            orientation="h",
            text="Haber Sayısı",
            title="İl Bazlı Yoğunluk"
        )
        fig.update_layout(
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            font_color="#e5e7eb",
            title_font_size=22,
            yaxis={"categoryorder": "total ascending"}
        )
        fig.update_traces(textfont_size=14)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("İl verisi yok.")

c3, c4 = st.columns([2, 1])

with c3:
    district_count = explode_count(filtered, "district_candidates", "İlçe", top_n=12)

    if not district_count.empty:
        fig = px.bar(
            district_count,
            x="İlçe",
            y="Haber Sayısı",
            text="Haber Sayısı",
            title="İlçe Dağılımı"
        )
        fig.update_layout(
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            font_color="#e5e7eb",
            title_font_size=22
        )
        fig.update_traces(textfont_size=14)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("İlçe verisi yok.")

with c4:
    st.markdown("### Mini Firma / Maps Listesi")
    firm_count = (
        filtered[filtered["firma_display"] != ""]["firma_display"]
        .value_counts()
        .head(8)
        .reset_index()
    )
    firm_count.columns = ["Firma / Maps", "Haber"]

    if not firm_count.empty:
        st.dataframe(firm_count, use_container_width=True, hide_index=True)
    else:
        st.info("Firma / Maps verisi yok.")

st.divider()

st.subheader("Kontrol Tablosu")

show_cols = [
    "published", "category", "title", "summary",
    "city_candidates", "district_candidates", "neighborhood_candidates",
    "firma_display", "activity_candidates", "link"
]

table_df = filtered[show_cols].copy()

table_df = table_df.rename(columns={
    "published": "Tarih",
    "category": "Kategori",
    "title": "Başlık",
    "summary": "Özet",
    "city_candidates": "İl",
    "district_candidates": "İlçe",
    "neighborhood_candidates": "Mahalle / Bölge",
    "firma_display": "Firma / Maps",
    "activity_candidates": "Faaliyet",
    "link": "Link"
})

table_df["Link"] = table_df["Link"].apply(
    lambda x: f'<a href="{x}" target="_blank">Habere Git</a>' if clean_text(x) else ""
)

st.write(table_df.to_html(escape=False, index=False), unsafe_allow_html=True)

st.divider()

st.subheader("Son Haber Kartları")

for _, row in filtered.head(60).iterrows():
    title = clean_text(row.get("title", ""))
    link = clean_text(row.get("link", ""))
    published = clean_text(row.get("published", ""))
    category = clean_text(row.get("category", ""))
    city = clean_text(row.get("city_candidates", ""))
    district = clean_text(row.get("district_candidates", ""))
    firm = clean_text(row.get("firma_display", ""))
    summary = clean_text(row.get("summary", ""))
    content = clean_text(row.get("content", ""))

    st.markdown('<div class="news-card">', unsafe_allow_html=True)

    st.markdown(f"### {title}")

    badges = []
    if category:
        badges.append(f"<span class='badge'>{category}</span>")
    if city:
        badges.append(f"<span class='badge'>{city}</span>")
    if district:
        badges.append(f"<span class='badge'>{district}</span>")
    if firm:
        badges.append(f"<span class='badge'>{firm}</span>")
    if published:
        badges.append(f"<span class='badge'>{published}</span>")

    if badges:
        st.markdown(" ".join(badges), unsafe_allow_html=True)

    if summary:
        st.write(summary)

    if content and content != summary:
        with st.expander("İçerik önizleme"):
            st.write(short_text(content, 1200))

    if link:
        st.link_button("Habere Git", link)

    st.markdown("</div>", unsafe_allow_html=True)

st.divider()

d1, d2 = st.columns(2)

with d1:
    st.download_button(
        "Filtreli CSV indir",
        data=filtered.to_csv(index=False).encode("utf-8-sig"),
        file_name="haberler_filtered.csv",
        mime="text/csv"
    )

with d2:
    st.download_button(
        "Tüm CSV indir",
        data=df.to_csv(index=False).encode("utf-8-sig"),
        file_name="haberler_all.csv",
        mime="text/csv"
    )