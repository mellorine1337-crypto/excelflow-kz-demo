from __future__ import annotations
import io
from pathlib import Path
import altair as alt
import pandas as pd
import streamlit as st
from analytics import FIELD_LABELS,FIELDS,excel_sheets,export_report,generate_insights,period_comparison,prepare_sales,product_abc,read_table,suggest_mapping,summarize
from database import save_run

st.set_page_config(page_title="ExcelFlow KZ",page_icon="📊",layout="wide",initial_sidebar_state="collapsed")
st.markdown("""<style>
:root{--navy:#0d2c47;--blue:#1877f2;--green:#22a06b;--ink:#172534;--muted:#66788a;--line:#dfe8f1}
.stApp{background:linear-gradient(180deg,#f7fbff 0,#fff 42%);color:var(--ink)}
.block-container{padding-top:1.15rem;padding-bottom:4rem;max-width:1480px}
.hero{position:relative;overflow:hidden;background:linear-gradient(125deg,#eef7ff 0%,#dceeff 58%,#cfe4ff 100%);border:1px solid #c4ddf6;border-radius:24px;padding:32px 36px;margin-bottom:22px;color:#15334d;box-shadow:0 16px 38px rgba(38,94,140,.12)}
.hero:after{content:"";position:absolute;width:280px;height:280px;border-radius:50%;right:-70px;top:-120px;background:rgba(24,119,242,.10)}
.hero-kicker{display:inline-block;padding:6px 11px;border:1px solid #7fb6e8;background:rgba(255,255,255,.6);color:#1769aa;border-radius:999px;font-size:.72rem;font-weight:800;letter-spacing:.12em}
.hero h1{font-size:2.35rem;line-height:1.05;margin:14px 0 8px;color:#12324d}.hero p{margin:0;color:#3d5b73;font-size:1.04rem;max-width:720px;font-weight:500}
.period-line{color:#607286;font-size:.88rem;font-weight:600;margin:2px 0 14px}
.section-head{margin:6px 0 12px}.section-head h2{font-size:1.35rem;margin:0;color:#17324a}.section-head p{color:#718294;margin:3px 0 0;font-size:.88rem}
[data-testid="stSidebar"]{background:linear-gradient(180deg,#f1f7fc 0%,#e8f2fa 100%);border-right:1px solid #d4e3ef}
[data-testid="stSidebar"] h2,[data-testid="stSidebar"] h3,[data-testid="stSidebar"] label,[data-testid="stSidebar"] p{color:#243f55}
[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"]{background:#fff;border-color:#bcd2e4}
[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] span,[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] small{color:#445e73!important}
[data-testid="stMetric"]{background:#fff;border:1px solid var(--line);border-top:4px solid var(--blue);border-radius:16px;padding:16px 17px;box-shadow:0 8px 24px rgba(36,72,101,.08)}
[data-testid="stMetricLabel"]{color:#6d7f91;font-weight:700}[data-testid="stMetricValue"]{color:#17324a;font-weight:800;font-size:clamp(1.35rem,1.75vw,1.85rem);letter-spacing:-.03em}[data-testid="stMetricValue"]>div{overflow:visible;text-overflow:clip}
div[data-testid="stAlert"]{border-radius:14px}.stButton button,.stDownloadButton button{min-height:3rem;border:0;border-radius:12px;background:linear-gradient(135deg,#1877f2,#1262c8);color:#fff;font-weight:800;box-shadow:0 7px 18px rgba(24,119,242,.20);transition:.18s ease}
.stButton button:hover,.stDownloadButton button:hover{color:#fff;border:0;transform:translateY(-1px);box-shadow:0 10px 24px rgba(24,119,242,.28)}
[data-testid="stSidebar"] .stButton button p,[data-testid="stSidebar"] .stButton button span{color:#fff!important}
div[data-baseweb="tab-list"]{gap:.45rem;background:#eaf1f7;padding:.42rem;border-radius:14px;margin:8px 0 18px}
button[data-baseweb="tab"]{border-radius:10px;padding:.55rem 1.15rem;color:#53687a;font-weight:700}button[data-baseweb="tab"][aria-selected="true"]{background:#fff;color:#1877f2;box-shadow:0 3px 10px rgba(36,72,101,.10)}
[data-testid="stDataFrame"]{border:1px solid var(--line);border-radius:14px;overflow:hidden}.insight-card{background:#fff;border:1px solid var(--line);border-left:4px solid #22a06b;border-radius:12px;padding:12px 14px;margin-bottom:9px;color:#33495d;box-shadow:0 5px 16px rgba(36,72,101,.05)}
.trust-card{background:#fff;border:1px solid var(--line);border-radius:18px;padding:24px;box-shadow:0 8px 24px rgba(36,72,101,.07)}
</style>""",unsafe_allow_html=True)
def money(v): return f"{v:,.0f} ₸".replace(","," ")
def delta(v,points=False): return None if v is None else f"{v:+.1f} {'п.п.' if points else '%'}"
def section_title(title,subtitle=""): st.markdown(f'<div class="section-head"><h2>{title}</h2><p>{subtitle}</p></div>',unsafe_allow_html=True)

st.markdown('''<div class="hero"><span class="hero-kicker">ЛОКАЛЬНАЯ БИЗНЕС-АНАЛИТИКА</span><h1>ExcelFlow KZ</h1><p>Профессиональный разбор продаж: прибыль, маржинальность, ассортимент и точки роста — из обычного Excel-файла.</p></div>''',unsafe_allow_html=True)
with st.sidebar:
    st.header("1. Данные"); uploaded=st.file_uploader("Загрузить Excel или CSV",type=["xlsx","csv"],help="Данные обрабатываются локально и не отправляются в облако."); demo=st.button("Вернуть демонстрационный отчёт",use_container_width=True); st.caption("Без загрузки автоматически показывается готовое демо. Для своего отчёта выберите файл до 10 МБ.")
using_demo=demo or uploaded is None
if using_demo:
    uploaded=io.BytesIO((Path(__file__).parent/"demo_sales.xlsx").read_bytes()); uploaded.name="demo_sales.xlsx"
try:
    with st.sidebar:
        sheets=excel_sheets(uploaded); sheet=st.selectbox("Лист Excel",sheets) if len(sheets)>1 else (sheets[0] if sheets else 0)
    raw=read_table(uploaded,sheet); suggested=suggest_mapping(raw.columns)
except ValueError as exc: st.error(str(exc)); st.stop()
with st.sidebar:
    st.header("2. Колонки")
    with st.expander("Проверить сопоставление",expanded=any(v is None for v in suggested.values())):
        options=["— не выбрано —",*raw.columns.tolist()]; mapping={}
        for field in FIELDS:
            proposed=suggested[field]; index=options.index(proposed) if proposed in options else 0
            selected=st.selectbox(FIELD_LABELS[field],options,index=index,key=f"map_{field}"); mapping[field]=None if selected==options[0] else selected
try: sales=prepare_sales(raw,mapping)
except ValueError as exc: st.error(str(exc)); st.caption("Откройте «Проверить сопоставление» слева и укажите нужные колонки."); st.stop()
with st.sidebar:
    st.header("3. Фильтры"); min_date,max_date=sales.date.min().date(),sales.date.max().date(); dates=st.date_input("Период",value=(min_date,max_date),min_value=min_date,max_value=max_date)
    if not isinstance(dates,(tuple,list)) or len(dates)!=2: st.warning("Выберите начало и конец периода."); st.stop()
    all_categories=sorted(sales["category"].dropna().unique().tolist()); categories=st.multiselect("Категории",all_categories,default=all_categories,placeholder="Выберите категории")
    category_sales=sales[sales["category"].isin(categories)]; all_products=sorted(category_sales["product"].dropna().unique().tolist()); products=st.multiselect("Товары",all_products,default=all_products,placeholder="Выберите товары"); compare=st.toggle("Сравнить с предыдущим периодом",value=True)
current,previous,changes=period_comparison(category_sales[category_sales["product"].isin(products)],dates[0],dates[1])
if current.empty: st.warning("За выбранный период и товары продаж нет. Измените фильтры."); st.stop()
if not compare: previous=previous.iloc[0:0]; changes={k:None for k in changes}
s=summarize(current); bp=current.groupby(["category","product"],as_index=False)[["revenue","gross_profit","quantity"]].sum(); bp["margin_pct"]=(bp.gross_profit/bp.revenue.where(bp.revenue!=0)*100).fillna(0); bp=bp.sort_values("revenue",ascending=False)
by_category=current.groupby("category",as_index=False)[["revenue","gross_profit","quantity"]].sum(); by_category["margin_pct"]=(by_category.gross_profit/by_category.revenue.where(by_category.revenue!=0)*100).fillna(0)
demo_label=" &nbsp;•&nbsp; Демонстрационные данные" if using_demo else ""
st.markdown(f'<div class="period-line">Период: {dates[0]:%d.%m.%Y} — {dates[1]:%d.%m.%Y} &nbsp;•&nbsp; {len(current):,} строк &nbsp;•&nbsp; {current["product"].nunique()} товаров{demo_label}</div>',unsafe_allow_html=True)
cols=st.columns(5); cols[0].metric("Выручка",money(s["revenue"]),delta(changes["revenue"])); cols[1].metric("Валовая прибыль",money(s["gross_profit"]),delta(changes["gross_profit"])); cols[2].metric("Валовая маржа",f"{s['margin_pct']:.1f}%",delta(changes["margin_pct"],True)); cols[3].metric("Продано",f"{s['units']:,.0f}".replace(","," "),delta(changes["units"])); cols[4].metric("Средний чек",money(s["avg_check"]),delta(changes["avg_check"]))
if compare and previous.empty: st.caption("ℹ️ В файле нет данных за предыдущий период такой же длины — дельты KPI не рассчитаны.")
overview,products_tab,abc_tab,data_tab=st.tabs(["Обзор","Товары","ABC-анализ","Данные"])
with overview:
    left,right=st.columns([1.8,1]); daily=current.assign(day=current.date.dt.normalize()).groupby("day",as_index=False)[["revenue","gross_profit"]].sum()
    with left:
        section_title("Динамика бизнеса","Как менялись выручка и валовая прибыль по дням")
        daily_long=daily.melt("day",var_name="metric",value_name="amount"); daily_long["metric"]=daily_long.metric.map({"revenue":"Выручка","gross_profit":"Валовая прибыль"})
        chart=alt.Chart(daily_long).mark_line(point=alt.OverlayMarkDef(filled=True,size=45),strokeWidth=3).encode(x=alt.X("day:T",title=None),y=alt.Y("amount:Q",title="Сумма, ₸"),color=alt.Color("metric:N",title=None,scale=alt.Scale(domain=["Выручка","Валовая прибыль"],range=["#1877F2","#22A06B"])),tooltip=[alt.Tooltip("day:T",title="Дата",format="%d.%m.%Y"),alt.Tooltip("metric:N",title="Показатель"),alt.Tooltip("amount:Q",title="Сумма",format=",.0f")]).properties(height=330); st.altair_chart(chart,use_container_width=True)
    with right:
        section_title("Выводы аналитика","Что заслуживает внимания прямо сейчас")
        for insight in generate_insights(current,previous if compare else None): st.markdown(f'<div class="insight-card">{insight}</div>',unsafe_allow_html=True)
    if compare and not previous.empty:
        comparison=pd.DataFrame({"Период":["Предыдущий","Текущий"],"Выручка":[summarize(previous)["revenue"],s["revenue"]],"Валовая прибыль":[summarize(previous)["gross_profit"],s["gross_profit"]]})
        section_title("Сравнение периодов","Текущий результат относительно предыдущего периода той же длины")
        st.altair_chart(alt.Chart(comparison.melt("Период",var_name="Показатель",value_name="Сумма")).mark_bar(cornerRadiusTopLeft=4,cornerRadiusTopRight=4).encode(x=alt.X("Период:N",title=None),y=alt.Y("Сумма:Q",title="Сумма, ₸"),color=alt.Color("Показатель:N",scale=alt.Scale(range=["#1877F2","#33A474"])),tooltip=["Период","Показатель",alt.Tooltip("Сумма:Q",format=",.0f")]).properties(height=260),use_container_width=True)
    ccat,cweek=st.columns(2); weekday=current.assign(weekday=current.date.dt.day_name().map({"Monday":"Пн","Tuesday":"Вт","Wednesday":"Ср","Thursday":"Чт","Friday":"Пт","Saturday":"Сб","Sunday":"Вс"})).groupby("weekday",as_index=False).revenue.sum()
    weekday["order"]=weekday.weekday.map({"Пн":1,"Вт":2,"Ср":3,"Чт":4,"Пт":5,"Сб":6,"Вс":7}); weekday=weekday.sort_values("order")
    with ccat:
        section_title("Структура выручки","Вклад категорий в общий результат")
        donut=alt.Chart(by_category).mark_arc(innerRadius=72,outerRadius=125,cornerRadius=4).encode(theta=alt.Theta("revenue:Q"),color=alt.Color("category:N",title="Категория",scale=alt.Scale(scheme="blues")),tooltip=[alt.Tooltip("category:N",title="Категория"),alt.Tooltip("revenue:Q",title="Выручка",format=",.0f"),alt.Tooltip("margin_pct:Q",title="Маржа",format=".1f")]).properties(height=285); st.altair_chart(donut,use_container_width=True)
    with cweek:
        section_title("Дни недели","Когда бизнес получает больше всего выручки")
        st.altair_chart(alt.Chart(weekday).mark_bar(color="#22A06B",cornerRadiusTopLeft=5,cornerRadiusTopRight=5).encode(x=alt.X("weekday:N",sort=weekday.weekday.tolist(),title=None),y=alt.Y("revenue:Q",title="Выручка, ₸"),tooltip=[alt.Tooltip("weekday:N",title="День"),alt.Tooltip("revenue:Q",title="Выручка",format=",.0f")]).properties(height=285),use_container_width=True)
with products_tab:
    section_title("Товарная аналитика","Лидеры продаж и баланс между оборотом и маржинальностью")
    pleft,pright=st.columns([1.15,1])
    with pleft:
        chart=alt.Chart(bp.head(15)).mark_bar(cornerRadiusEnd=5,color="#1877F2").encode(y=alt.Y("product:N",sort="-x",title=None),x=alt.X("revenue:Q",title="Выручка, ₸"),tooltip=[alt.Tooltip("product:N",title="Товар"),alt.Tooltip("revenue:Q",title="Выручка",format=",.0f"),alt.Tooltip("margin_pct:Q",title="Маржа",format=".1f")]).properties(height=420); st.altair_chart(chart,use_container_width=True)
    with pright:
        scatter=alt.Chart(bp.head(40)).mark_circle(opacity=.78,stroke="white",strokeWidth=1).encode(x=alt.X("revenue:Q",title="Выручка, ₸"),y=alt.Y("margin_pct:Q",title="Маржа, %"),size=alt.Size("quantity:Q",title="Количество",scale=alt.Scale(range=[80,1300])),color=alt.Color("category:N",title="Категория",scale=alt.Scale(scheme="tableau10")),tooltip=[alt.Tooltip("product:N",title="Товар"),alt.Tooltip("category:N",title="Категория"),alt.Tooltip("revenue:Q",title="Выручка",format=",.0f"),alt.Tooltip("gross_profit:Q",title="Прибыль",format=",.0f"),alt.Tooltip("margin_pct:Q",title="Маржа",format=".1f")]).properties(height=420); zero=alt.Chart(pd.DataFrame({"y":[0]})).mark_rule(color="#dc5a5a",strokeDash=[6,4]).encode(y="y:Q"); st.altair_chart(scatter+zero,use_container_width=True)
    st.dataframe(bp.rename(columns={"category":"Категория","product":"Товар","revenue":"Выручка","gross_profit":"Валовая прибыль","quantity":"Количество","margin_pct":"Маржа, %"}),use_container_width=True,hide_index=True,column_config={"Выручка":st.column_config.NumberColumn(format="%.0f ₸"),"Валовая прибыль":st.column_config.NumberColumn(format="%.0f ₸"),"Маржа, %":st.column_config.ProgressColumn(min_value=-100,max_value=100,format="%.1f%%")})
with abc_tab:
    section_title("ABC-анализ ассортимента","Класс A формирует примерно 80% выручки, B — следующие 15%, C — оставшуюся часть")
    abc_raw=product_abc(current); abc_summary=abc_raw.groupby("abc_class",as_index=False).agg(revenue=("revenue","sum"),products=("product","nunique"))
    ac1,ac2=st.columns(2)
    with ac1: st.altair_chart(alt.Chart(abc_summary).mark_bar(cornerRadiusTopLeft=6,cornerRadiusTopRight=6).encode(x=alt.X("abc_class:N",title="Класс",sort=["A","B","C"]),y=alt.Y("revenue:Q",title="Выручка, ₸"),color=alt.Color("abc_class:N",title=None,scale=alt.Scale(domain=["A","B","C"],range=["#1877F2","#22A06B","#F3A83B"])),tooltip=[alt.Tooltip("abc_class:N",title="Класс"),alt.Tooltip("revenue:Q",title="Выручка",format=",.0f"),alt.Tooltip("products:Q",title="Товаров")]).properties(height=260),use_container_width=True)
    with ac2: st.altair_chart(alt.Chart(abc_raw).mark_line(point=True,color="#1877F2",strokeWidth=3).encode(x=alt.X("product:N",sort=None,title="Товары от лидеров к остальным",axis=alt.Axis(labels=False,ticks=False)),y=alt.Y("cumulative_share_pct:Q",title="Накопленная доля, %",scale=alt.Scale(domain=[0,100])),tooltip=[alt.Tooltip("product:N",title="Товар"),alt.Tooltip("cumulative_share_pct:Q",title="Накопленная доля",format=".1f"),alt.Tooltip("abc_class:N",title="Класс")]).properties(height=260),use_container_width=True)
    abc=abc_raw.rename(columns={"product":"Товар","revenue":"Выручка","gross_profit":"Валовая прибыль","quantity":"Количество","revenue_share_pct":"Доля выручки, %","cumulative_share_pct":"Накопленная доля, %","abc_class":"Класс"})
    st.dataframe(abc,use_container_width=True,hide_index=True,column_config={"Выручка":st.column_config.NumberColumn(format="%.0f ₸"),"Валовая прибыль":st.column_config.NumberColumn(format="%.0f ₸"),"Доля выручки, %":st.column_config.ProgressColumn(min_value=0,max_value=100,format="%.1f%%"),"Накопленная доля, %":st.column_config.ProgressColumn(min_value=0,max_value=100,format="%.1f%%")})
with data_tab:
    display=current.rename(columns={"date":"Дата","category":"Категория","product":"Товар","quantity":"Количество","price":"Цена","cost":"Себестоимость","revenue":"Выручка","gross_profit":"Валовая прибыль","margin_pct":"Маржа, %"}); st.dataframe(display,use_container_width=True,hide_index=True)
st.divider(); report=export_report(current,previous if compare else None); c1,c2,_=st.columns([1,1,2]); c1.download_button("Скачать Excel-отчёт",data=report,file_name=f"ExcelFlow_{dates[0]:%Y%m%d}_{dates[1]:%Y%m%d}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
if c2.button("Сохранить итоги в историю",use_container_width=True): save_run(s,len(current)); st.toast("Итоги сохранены локально",icon="✅")
