import pandas as pd
import numpy as np


def clean_text_column(series):
    """Сверхточная зачистка текстовых данных от скрытых пробелов"""
    return series.fillna("").astype(str).str.strip().str.replace(r'\s+', ' ', regex=True)


def process_data(stock_path, sales_path):
    """Загрузка сырых данных, очистка, расчет метрик и скоринг оборачиваемости"""
    df_stock = pd.read_excel(stock_path)
    df_sales = pd.read_excel(sales_path)

    df_stock.columns = df_stock.columns.str.strip()
    df_sales.columns = df_sales.columns.str.strip()

    sales_item_col = df_sales.columns[0]

    st_df = df_stock.iloc[:, :3].copy()
    st_df.columns = ["Товар", "Остаток", "Дней на складе"]

    df_sales["D_Obj"] = pd.to_datetime(df_sales.iloc[:, 1], errors="coerce")

    # ИСПРАВЛЕНИЕ: Фиксируем точные границы отрезка времени прямо из данных файла продаж
    DATA_START_DATE = df_sales["D_Obj"].min()
    GLOBAL_LAST_DATE = df_sales["D_Obj"].max()

    def agg_sales(g):
        sd = g.dropna().sort_values()
        cnt = len(sd)
        txt = ", ".join(sd.dt.strftime("%d.%m.%Y").unique())

        # ИСПРАВЛЕНИЕ: Точка отсчета жестко привязана к дате последней известной операции
        CURRENT_DATE = GLOBAL_LAST_DATE

        if cnt > 0:
            # Период считаем в рамках фиксированного цикла времени (от Дня А до Дня Д)
            total_days_active = max(
                1, (CURRENT_DATE - DATA_START_DATE).days + 1)
            sales_per_day = cnt / total_days_active

            # Дней между первой и последней фактической продажей товара внутри периода
            if cnt > 1:
                days_range = max(1, (sd.max() - sd.min()).days)
                mean_interval = days_range / cnt
            else:
                mean_interval = 999999.0

            # Сколько дней товар молчал до окончания отчетного периода (Дня Д)
            days_since_last_sale = (CURRENT_DATE - sd.max()).days
        else:
            sales_per_day = 0.0
            mean_interval = 999999.0
            days_since_last_sale = (CURRENT_DATE - DATA_START_DATE).days + 1

        return pd.Series([cnt, txt, mean_interval, sales_per_day, days_since_last_sale],
                         index=["P", "C", "I", "S", "Days_Silent"])

    sl_df = df_sales.groupby(sales_item_col)["D_Obj"].apply(
        agg_sales).unstack().reset_index()
    sl_df.rename(columns={sales_item_col: "Товар"}, inplace=True)

    st_df["Товар"] = clean_text_column(st_df["Товар"])
    sl_df["Товар"] = clean_text_column(sl_df["Товар"])

    df = pd.merge(st_df, sl_df, on="Товар", how="outer")
    df["P"] = df["P"].fillna(0).astype(int)
    df["Остаток"] = df["Остаток"].fillna(0).astype(int)
    df["Дней на складе"] = df["Дней на складе"].fillna("Нет данных")
    df["C"] = df["C"].fillna("Нет продаж")
    df["S"] = df["S"].fillna(0.0)
    df["Days_Silent"] = df["Days_Silent"].fillna(999999).astype(int)
    df["I"] = df["I"].fillna(999999.0).astype(float)

    # --- МАТЕМАТИЧЕСКИЙ СКОРИНГ ОБОРОТА ---
    turnover_labels = []
    for idx, row in df.iterrows():
        p_count = row["P"]
        speed = row["S"]
        silent = row["Days_Silent"]

        if p_count == 0 or silent > 90:
            turnover_labels.append("1-низкий")
        elif speed >= 0.10 and silent <= 30:
            turnover_labels.append("3-высокий")
        elif speed >= 0.033 and silent <= 45:
            turnover_labels.append("2-средний")
        else:
            turnover_labels.append("1-низкий")

    df["O"] = turnover_labels
    return df


def segment_into_groups(df):
    """Распределение позиций по пулам спроса A, B, C и выделение HiPo"""
    live_demand_pool = df[(df["P"] >= 2) & (df["Days_Silent"] <= 45)]
    hipo = live_demand_pool.sort_values(
        by=["I", "P"], ascending=[True, False]).head(20).copy()

    remaining_df = df.drop(hipo.index, errors='ignore')

    pool_A = remaining_df[remaining_df["P"] >= 3].copy()
    pool_B = remaining_df[remaining_df["P"].isin([1, 2])].copy()
    pool_C = remaining_df[remaining_df["P"] == 0].copy()

    combined_A_pool = pd.concat(
        [pool_A, hipo]).drop_duplicates(subset=["Товар"]).copy()

    t5 = combined_A_pool.sort_values(
        ["P", "Остаток"], ascending=[False, True]).head(5)
    rem = combined_A_pool.drop(t5.index)
    t10 = rem.sort_values(["P", "Остаток"], ascending=[False, True]).head(5)
    rem = rem.drop(t10.index)

    actual_hipo = hipo.drop(t5.index, errors='ignore').drop(
        t10.index, errors='ignore')
    tail_A = rem.drop(actual_hipo.index, errors='ignore').sort_values(
        ["P", "Остаток"], ascending=[False, True])

    fin_A = pd.concat([t5, t10, actual_hipo, tail_A]).reset_index(drop=True)
    fin_B = pool_B.sort_values(["P", "Остаток"], ascending=[
                               False, True]).reset_index(drop=True)
    fin_C = pool_C.sort_values(["Остаток", "Дней на складе"], ascending=[
                               False, False]).reset_index(drop=True)

    return fin_A, fin_B, fin_C, actual_hipo
