import pandas as pd
from wp_recipes import new_recipes
from wp_weights import material_weights

# === 使用者輸入 ===
max_weight = float(input("請輸入你的負重上限（公斤）："))
boost = float(input("負重上限提升幾%：")) / 100
RESOURCE_RETURN_RATE = float(input("資源返還率是多少%：")) / 100
vip_input = input("是否有VIP? Y/N：").strip().upper()
vip_discount = 0.04 if vip_input == "Y" else 0.08

effective_weight_limit = max_weight * (1 + boost)

# === 讀取資料 ===
wp_df = pd.read_excel('裝備價格.xlsx')
make_wp_df = pd.read_excel('製作素材價格.xlsx')

# === 城市名 ===
cities = ["Martlock", "Thetford", "Bridgewatch", "Lymhurst", "Fort Sterling", "Brecilien"]
all_cities = cities + ["Caerleon"]

# === 開始計算 ===
profit_results = []
weight_results = []

for _, row in wp_df.iterrows():
    item_name = row["物品名稱"]
    enchant = int(row["附魔"]) if not pd.isna(row["附魔"]) else 0
    if item_name not in new_recipes:
        continue

    recipe = new_recipes[item_name]
    material_cost = 0
    total_weight = 0
    material_counts = []

    for material, qty in recipe.items():
        # 成本（考慮返還率）
        mat_row = make_wp_df[
            (make_wp_df["物品名稱"] == material) &
            (make_wp_df["附魔"] == enchant)
        ]
        if mat_row.empty:
            min_price = 0
        else:
            prices = mat_row[cities].values[0]
            min_price = min([p for p in prices if p > 0], default=0)

        adjusted_qty = qty * (1 - RESOURCE_RETURN_RATE)
        material_cost += min_price * adjusted_qty

        # 負重（不考慮返還率）
        weight_per_unit = material_weights.get(material, 0)
        total_weight += weight_per_unit * qty
        material_counts.append((material, qty))

    row_result = {
        "物品名稱": item_name,
        "附魔": enchant,
        "單件成本": round(material_cost, 2) if material_cost > 0 else None
    }

    for city in all_cities:
        sell_price = row.get(city, 0)
        if material_cost > 0 and sell_price > 0:
            # 利潤計算（考慮稅率與設定費）
            true_profit = sell_price * (1 - vip_discount - 0.025) - material_cost
            row_result[city] = round(true_profit, 2)
        else:
            row_result[city] = None

    profit_results.append(row_result)

    # === 計算可製作幾件（不考慮返還率）===
    if total_weight > 0:
        max_units = int(effective_weight_limit // total_weight)
        weight_row = {
            "裝備名稱": item_name,
            "附魔": enchant,
            "可以做幾件": max_units
        }
        for i, (mat, qty) in enumerate(material_counts, start=1):
            weight_row[f"素材{i}"] = mat
            weight_row[f"件數{i}"] = qty * max_units
        weight_results.append(weight_row)

# === 輸出 Excel：裝備利潤分析 + 負重分配表 ===
profit_df = pd.DataFrame(profit_results)
profit_df = profit_df[["物品名稱", "附魔", "單件成本"] + all_cities]

weight_df = pd.DataFrame(weight_results)

with pd.ExcelWriter("裝備利潤分析.xlsx", engine="openpyxl", mode="w") as writer:
    profit_df.to_excel(writer, sheet_name="裝備利潤分析", index=False)
    weight_df.to_excel(writer, sheet_name="負重分配表", index=False)

# === 輸出使用到的素材價格 ===
all_materials = set()
for recipe in new_recipes.values():
    all_materials.update(recipe.keys())

filtered_df = make_wp_df[make_wp_df["物品名稱"].isin(all_materials)].copy()
columns_order = ["物品名稱", "附魔"] + all_cities
filtered_df = filtered_df[columns_order]
filtered_df.to_excel("使用到的製作素材價格.xlsx", index=False)
