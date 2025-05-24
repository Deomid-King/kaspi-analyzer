import streamlit as st
import pandas as pd
from io import BytesIO
import matplotlib.pyplot as plt

st.set_page_config(page_title="Kaspi Анализ Заказов", layout="wide")
st.title("📦 Kaspi Анализ заказов с расчетом прибыли")

if "costs" not in st.session_state:
    st.session_state.costs = {}
if "search_query" not in st.session_state:
    st.session_state.search_query = ""
if "sort_by" not in st.session_state:
    st.session_state.sort_by = "Кол_заказов"

@st.cache_data
def load_data(uploaded_file):
    df = pd.read_excel(uploaded_file, sheet_name=0)
    df = df[[
        "Дата поступления заказа",
        "Название в системе продавца",
        "Артикул",
        "Сумма",
        "Статус",
        "Количество",
        "Стоимость доставки для продавца",
        "Склад передачи КД"
    ]].copy()
    df.columns = [
        "Дата заказа",
        "Название товара",
        "Артикул",
        "Сумма",
        "Статус",
        "Количество",
        "Доставка",
        "Склад"
    ]
    df["Дата заказа"] = pd.to_datetime(df["Дата заказа"], dayfirst=True, errors='coerce')
    warehouse_names = {
        "2667005_PP1": "Петропавловск",
        "2667005_PP24": "Алматы"
    }
    df["Склад"] = df["Склад"].replace(warehouse_names)
    return df

uploaded_file = st.file_uploader("Загрузите Excel-файл с заказами", type=["xlsx"])

if uploaded_file:
    df = load_data(uploaded_file)

    st.sidebar.header("🔍 Фильтры")
    min_date = df["Дата заказа"].min()
    max_date = df["Дата заказа"].max()
    date_range = st.sidebar.date_input("Период заказов", [min_date, max_date])

    returns_df = df[
        (df["Дата заказа"] >= pd.to_datetime(date_range[0])) &
        (df["Дата заказа"] <= pd.to_datetime(date_range[1])) &
        (df["Статус"] == "Возврат")
    ]

    total_sales = df[
        (df["Дата заказа"] >= pd.to_datetime(date_range[0])) &
        (df["Дата заказа"] <= pd.to_datetime(date_range[1])) &
        (df["Статус"] == "Выдан")
    ]["Сумма"].sum()

    html_stats = f"""
    <div style='background-color: #f5f5f5; padding: 1em; border-radius: 8px;'>
        <p style='margin-bottom: 0.5em;'>🔁 <strong>Возвратов:</strong> {len(returns_df)} заказов на сумму <strong>{returns_df['Сумма'].sum():,.0f} ₸</strong></p>
        <p style='margin-bottom: 0; font-size: 1.2em;'><strong>💵 Оборот:</strong> {total_sales:,.0f} ₸</p>
    </div>
    """
    st.sidebar.markdown(html_stats, unsafe_allow_html=True)

    issued_df = df[
        (df["Дата заказа"] >= pd.to_datetime(date_range[0])) &
        (df["Дата заказа"] <= pd.to_datetime(date_range[1])) &
        (df["Статус"] == "Выдан")
    ]

    warehouses = issued_df["Склад"].dropna().unique().tolist()
    selected_warehouses = st.sidebar.multiselect("Склады", warehouses, default=warehouses)
    filtered_df = issued_df[issued_df["Склад"].isin(selected_warehouses)]

    st.sidebar.header("💰 Себестоимость")
    st.sidebar.text_input("🔎 Поиск по артикулу или названию", key="search_query")
    unique_products = filtered_df[["Артикул", "Название товара"]].drop_duplicates()
    search_query = st.session_state.search_query.lower().strip()
    if search_query:
        unique_products = unique_products[
            unique_products["Артикул"].str.lower().str.contains(search_query) |
            unique_products["Название товара"].str.lower().str.contains(search_query)
        ]
    for i, (_, row) in enumerate(unique_products.iterrows()):
        art = row["Артикул"]
        name = row["Название товара"]
        default_cost = st.session_state.costs.get(art, 0.0)
        cost = st.sidebar.number_input(f"{name} ({art})", min_value=0.0, value=default_cost, step=100.0, key=f"cost_{art}_{i}")
        st.session_state.costs[art] = cost

    filtered_df["Себестоимость"] = filtered_df["Артикул"].apply(lambda art: st.session_state.costs.get(art, 0.0))
    filtered_df["Чистая маржа"] = (
        (filtered_df["Сумма"] * 0.83)
        - (filtered_df["Себестоимость"] * filtered_df["Количество"])
        - (filtered_df["Доставка"] * filtered_df["Количество"])
    )

    summary = filtered_df.groupby(["Артикул", "Название товара", "Склад"]).agg(
        Кол_заказов=("Количество", "sum"),
        Сумма_продаж=("Сумма", "sum"),
        Средняя_себестоимость=("Себестоимость", "mean"),
        Общая_маржа=("Чистая маржа", "sum")
    ).reset_index()

    summary = summary.sort_values(by="Кол_заказов", ascending=False)
    st.subheader("📋 Сводка по заказам")
    st.dataframe(summary, use_container_width=True)

    st.markdown("## 🌟 Топ-10 товаров")
    sort_option = st.radio("Сортировка по:", ["Кол_заказов", "Сумма_продаж", "Общая_маржа"], horizontal=True)
    top10 = summary.sort_values(by=sort_option, ascending=False).head(10)
    st.dataframe(top10, use_container_width=True)

    with st.expander("📊 Диаграмма распределения заказов по складам"):
        pie_data = filtered_df.groupby("Склад")["Количество"].sum()
        fig, ax = plt.subplots()
        ax.pie(pie_data, labels=pie_data.index, autopct="%1.1f%%", startangle=90)
        ax.axis("equal")
        st.pyplot(fig)

    to_excel = BytesIO()
    with pd.ExcelWriter(to_excel, engine="xlsxwriter") as writer:
        summary.to_excel(writer, index=False, sheet_name="Сводка")
        filtered_df.to_excel(writer, index=False, sheet_name="Детали")
    to_excel.seek(0)
    st.download_button(
        label="📥 Скачать отчет в Excel",
        data=to_excel,
        file_name="kaspi_otchet.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
