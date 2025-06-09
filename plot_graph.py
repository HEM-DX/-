#"C:\Users\J0134011\OneDrive - Honda\デスクトップ\シーラー管理\HEM K-40改.xlsx" ファイル名

#横向き → x に数値列、y にカテゴリ列
#縦向き → x にカテゴリ列、y に数値列

#箱ひげ図	"box"	
#バイオリンプロット	"violin"
#棒グラフ	"bar"
#ポイントプロット	"point"	
#スウォームプロット	"swarm"
#ストリッププロット	"strip"	
#全体の分布 + 外れ値の確認 → "box" or "violin"
#個別データを全部見たい → "swarm" or "strip"
#平均値重視で見せたい → "bar" or "point"


import pandas as pd
import math
import seaborn as sns
import matplotlib.pyplot as plt

# ======= ファイル読み込み & 使用量数値変換 =======
file_path = r"C:\Users\J0134011\OneDrive - Honda\デスクトップ\シーラー管理\HEM K-40改.xlsx"
df = pd.read_excel(file_path)

# 単位 "g" を削除して数値に変換
df["使用量(g)"] = pd.to_numeric(df["使用量(g)"].astype(str).str.replace("g", "", regex=False), errors="coerce")
df = df.dropna(subset=["使用量(g)"])
df = df[df["使用量(g)"] > 0]

# ======= 各種パラメータ =======
production_per_day = 800    # 生産計画（台/日）
operating_days = 20         # 稼働日数（日）
drum_weight = 250_000  # ドラム缶1本の容量（g） ← 250kg


total_production = production_per_day * operating_days

# ======= 集計（平均→総使用量・必要本数） =======
group_cols = ["工程", "機種", "材質", "R/B"]
summary = df.groupby(group_cols)["使用量(g)"].mean().reset_index()
summary["総使用量[g]"] = summary["使用量(g)"] * total_production
summary["必要本数"] = summary["総使用量[g]"] / drum_weight
summary["必要本数"] = summary["必要本数"].apply(lambda x: math.ceil(x))

# ======= 結果出力（平均使用量は出さない） =======
print(summary[["工程", "機種", "材質", "R/B", "総使用量[g]", "必要本数"]])

# ======= グラフ：必要本数のみ表示 =======
sns.set(style="whitegrid", font="Meiryo")

g = sns.catplot(
    data=summary,
    kind="bar",
    x="機種",
    y="使用量(g)",
    hue="R/B",
    col="材質",
    row="工程",
    palette="pastel",
    height=5,
    aspect=1.5,
    ci=None
)

g.set_axis_labels("機種", "使用量(g)")
g.set_titles(col_template="{col_name}", row_template="{row_name}")
g.fig.subplots_adjust(top=0.9)
g.fig.suptitle("工程・材質・部位別 機種ごとの使用量", fontsize=16)

# ======= ラベル：必要本数だけ表示 =======
for ax, ((工程, 材質), subdata) in zip(g.axes.flat, summary.groupby(["工程", "材質"])):
    xticks = [t.get_text() for t in ax.get_xticklabels()]
    for i, (rb, rb_group) in enumerate(subdata.groupby("R/B")):
        for j, (機種名, g2) in enumerate(rb_group.groupby("機種")):
            y_val = g2["使用量(g)"].values[0]
            need_num = g2["必要本数"].values[0]
            if 機種名 in xticks:
                x_pos = xticks.index(機種名)
                offset = 0.15 * i
                ax.text(x_pos, y_val + offset, f"{need_num} 本", ha="center", fontsize=8, color="blue")

plt.show()
input("グラフを閉じるには Enter キーを押してください...")

# ======= 特定条件（材質がK-40 または E51G-JP または 1085G）の出力 =======
target_df = df[
    (df["材質"] == "K-40") |
    (df["材質"] == "E51G-JP") |
    (df["材質"] == "1085G")
]

if not target_df.empty:
    mean_val = target_df["使用量(g)"].mean()
    total_usage = mean_val * total_production
    total_usage_kg = total_usage / 1000
    need_drums = math.ceil(total_usage / drum_weight)

    print("\n===【材質がK-40 または E51G-JP または 1085G】の集計結果 ===")
    print(f"総使用量: {total_usage:,.0f} g（= {total_usage_kg:.1f} kg）")
    print(f"必要本数: {need_drums} 本")
else:
    print("\n指定の条件のデータが見つかりませんでした。")

# ======= 条件ごとの総使用量と必要本数（個別集計） =======
conditions = {
    "工程: K-40": (df["工程"] == "K-40"),
    "機種: E51G-JP": (df["機種"] == "E51G-JP"),
    "材質: K-40": (df["材質"] == "K-40"),
    "材質: E51G-JP": (df["材質"] == "E51G-JP"),
    "材質: 1085G": (df["材質"] == "1085G"),
}

print("\n===【機種 × 材質 ごとの集計結果】===\n")

grouped = df.groupby(["機種", "材質"])["使用量(g)"].mean().reset_index()

for _, row in grouped.iterrows():
    機種名 = row["機種"]
    材質名 = row["材質"]
    平均使用量 = row["使用量(g)"]

    total_usage = 平均使用量 * total_production
    total_usage_kg = total_usage / 1000
    need_drums = math.ceil(total_usage / drum_weight)

    print(f"{機種名} × {材質名}")
    print(f"　総使用量: {total_usage:,.0f} g（= {total_usage_kg:.1f} kg）")
    print(f"　必要本数: {need_drums} 本\n")



print("\n===【個別条件ごとの集計結果】===\n")

for label, condition in conditions.items():
    filtered_df = df[condition]
    if not filtered_df.empty:
        mean_val = filtered_df["使用量(g)"].mean()
        total_usage = mean_val * total_production
        total_usage_kg = total_usage / 1000
        need_drums = math.ceil(total_usage / drum_weight)

        print(f"{label}")
        print(f"　総使用量: {total_usage:,.0f} g（= {total_usage_kg:.1f} kg）")
        print(f"　必要本数: {need_drums} 本\n")
    else:
        print(f"{label} のデータが見つかりませんでした。\n")

        import streamlit as st

# ==== Streamlit 表示 ====
st.title("機種 × 材質ごとの合計使用量と必要本数")

# データテーブルの表示
st.dataframe(combo_group[["機種", "材質", "総使用量[g]", "必要本数"]])

# オプションで CSV ダウンロードも追加
csv = combo_group.to_csv(index=False).encode('utf-8')
st.download_button(
    label="CSVとしてダウンロード",
    data=csv,
    file_name='機種_材質_使用量(g)_必要本数.csv',
    mime='text/csv',
)


