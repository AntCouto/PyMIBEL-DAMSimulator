import pandas as pd
import pypsa
import numpy as np
import matplotlib.pyplot as plt

np.random.seed(42)

# =========================================================
# 1. LOAD DATA
# =========================================================
file_path = "Input_Exemplo.xlsx"
df = pd.read_excel(file_path)

df.columns = df.columns.str.strip().str.upper()
for c in ["TRANSACTION TYPE", "COUNTRY", "TECHNOLOGY", "UNIT"]:
    if c in df.columns:
        df[c] = df[c].astype(str).str.strip().str.upper()

# Alias de preço
df["BID PRICE (EUR/MWH)"] = df["BID PRICE RANDOM LNEG 2 (EUR/MWH)"]
# Artefact - small random perturbation to break price ties (NOT pro-rata)
df["BID PRICE (EUR/MWH)"] = df["BID PRICE (EUR/MWH)"] + 0.001 * np.random.rand(len(df))

# =========================================================
# 2. HOURS TO SIMULATE
# =========================================================
cols_time = ["PERIOD OF YEAR", "YEAR", "MONTH", "DAY", "SESSION", "PERIOD"]

hours = (
    df[cols_time]
    .dropna(subset=["PERIOD"])
    .drop_duplicates(subset=["PERIOD"])
    .sort_values("PERIOD")
    .reset_index(drop=True)
)
hours[cols_time] = hours[cols_time].apply(pd.to_numeric, errors="coerce").astype(int)
hours = hours.dropna(subset=cols_time)

# =========================================================
# 3. RESULTS CONTAINERS
# =========================================================
session_results = []
trading_results = []

# =========================================================
# 4. LOOP PER HOUR (DAY-AHEAD CLEARING)
# =========================================================
for _, h in hours.iterrows():

    period = h["PERIOD"]
    print(
        f"\n===== SIMULATION HOUR {period} =====\n"
        f"PERIOD OF YEAR: {h['PERIOD OF YEAR']}\n"
        f"Date: {h['YEAR']}-{h['MONTH']:02d}-{h['DAY']:02d}\n"
        f"SESSION: {h['SESSION']} | PERIOD: {period}\n"
    )

    # Filter bids of THIS hour only
    df_h = df[df["PERIOD"] == period].copy()

    # Split bids
    pt_sell = df_h[(df_h["COUNTRY"] == "PT") & (df_h["TRANSACTION TYPE"] == "SELL")]
    es_sell = df_h[(df_h["COUNTRY"] == "ES") & (df_h["TRANSACTION TYPE"] == "SELL")]
    pt_buy  = df_h[(df_h["COUNTRY"] == "PT") & (df_h["TRANSACTION TYPE"] == "BUY")]
    es_buy  = df_h[(df_h["COUNTRY"] == "ES") & (df_h["TRANSACTION TYPE"] == "BUY")]

    # -------------------------------
    # Create network for THIS hour
    # -------------------------------
    n = pypsa.Network()
    n.set_snapshots([0])
    n.add("Carrier", "electricity")
    n.add("Bus", "PT", carrier="electricity")
    n.add("Bus", "ES", carrier="electricity")

    # Interconnection capacity (hour-specific)
    intercap = df_h["INTERCONNECTION"].iloc[0] if "INTERCONNECTION" in df_h.columns else 3800
    n.add(
        "Link",
        "PT_ES",
        bus0="PT",
        bus1="ES",
        p_nom=float(intercap),
        p_min_pu=-1,
        carrier="electricity"
    )

    # -------------------------------
    # SELL BIDS
    # -------------------------------
    def add_sell(df_zone, zone):
        for i, r in df_zone.iterrows():
            e = float(r["BID ENERGY (MWH)"])
            if e <= 0:
                continue
            n.add(
                "Generator",
                name=f"SELL_{zone}_{i}",
                bus=zone,
                p_nom=e,
                marginal_cost=float(r["BID PRICE (EUR/MWH)"]),
                carrier="electricity"
            )

    add_sell(pt_sell, "PT")
    add_sell(es_sell, "ES")

    # -------------------------------
    # BUY BIDS (MANDATORY + FLEX)
    # -------------------------------
    MANDATORY_PRICE = 3880

    def add_buy(df_zone, zone):
        for i, r in df_zone.iterrows():
            e = float(r["BID ENERGY (MWH)"])
            if e <= 0:
                continue
            price = float(r["BID PRICE (EUR/MWH)"])
            if price >= MANDATORY_PRICE:
                n.add(
                    "Load",
                    name=f"BUY_{zone}_{i}",
                    bus=zone,
                    p_set=e
                )
            else:
                n.add(
                    "Generator",
                    name=f"FLEX_{zone}_{i}",
                    bus=zone,
                    p_nom=e,
                    p_min_pu=-1,
                    p_max_pu=0,
                    marginal_cost=price
                )

    add_buy(pt_buy, "PT")
    add_buy(es_buy, "ES")

    # -------------------------------
    # Solve market
    # -------------------------------
    n.optimize(solver_name="glpk")

    # -------------------------------
    # Prices
    # -------------------------------
    price_pt = n.buses_t.marginal_price.loc[0, "PT"]
    price_es = n.buses_t.marginal_price.loc[0, "ES"]
    clearing_price = max(price_pt, price_es)

    # -------------------------------
    # Interconnection flow (both directions)
    # -------------------------------
    link_flow = n.links_t.p0.loc[0, "PT_ES"] if "PT_ES" in n.links.index else 0
    pt_to_es_flow = max(link_flow, 0)
    es_to_pt_flow = max(-link_flow, 0)

    print(f"\n>>> INTERCONNECTION PT–ES")
    print(f"  Capacity (p_nom): {intercap} MW")
    print(f"  Flow PT→ES: {pt_to_es_flow:.2f} MW, ES→PT: {es_to_pt_flow:.2f} MW ({'PT → ES' if link_flow>0 else 'ES → PT' if link_flow<0 else '0'})")

    # -------------------------------
    # Print per zone
    # -------------------------------
    for zone in ["PT", "ES"]:
        price = n.buses_t.marginal_price.loc[0, zone]
        gens = n.generators[n.generators.bus == zone]
        sell = gens[~gens.index.str.contains("FLEX")].index
        flex = gens[gens.index.str.contains("FLEX")].index

        gen_sum = n.generators_t.p.loc[0, sell].sum() if len(sell) > 0 else 0
        flex_sum = abs(n.generators_t.p.loc[0, flex].sum()) if len(flex) > 0 else 0
        load_sum = n.loads_t.p.loc[0, n.loads.bus == zone].sum() if len(n.loads_t.p.loc[0, n.loads.bus == zone]) > 0 else 0

        print(f"\n>>> ZONE: {zone} | Price: {price:.2f} EUR/MWh")
        print(f"  (+) Supply (SELL):      {gen_sum:10.2f} MW")
        print(f"  [-] Flexible Demand:    {flex_sum:10.2f} MW")
        print(f"  [-] Mandatory Demand:   {load_sum:10.2f} MW")

    # -------------------------------
    # Aggregates
    # -------------------------------
    total_demand = n.loads_t.p.loc[0].sum()
    total_supply = n.generators_t.p.loc[0][n.generators.p_nom > 0].sum()
    traded = min(total_demand, total_supply)

    # Last cleared units
    gen_p = n.generators_t.p.loc[0]
    last_supply = gen_p[gen_p > 0].idxmax() if any(gen_p > 0) else ""
    last_demand = gen_p[gen_p < 0].idxmin() if any(gen_p < 0) else ""

    # -------------------------------
    # Session Results (1 ROW PER HOUR)
    # -------------------------------
    session_results.append({
        **h.to_dict(),
        "Bidding Area": "MI",
        "Pool Result": "TRADING",
        "Price_PT (EUR/MWh)": price_pt,
        "Price_ES (EUR/MWh)": price_es,
        "Clearing Price (EUR/MWh)": clearing_price,
        "PT→ES Flow (MW)": pt_to_es_flow,
        "ES→PT Flow (MW)": es_to_pt_flow,
        "Total Demand (MWh)": total_demand,
        "Total Supply (MWh)": total_supply,
        "Total Traded Energy (MWh)": traded,
        "Last Demand Trading Unit": last_demand,
        "Last Supply Trading Unit": last_supply
    })

    # -------------------------------
    # Trading Results Detailed
    # -------------------------------
    for i, r in df_h.iterrows():
        unit_name = f"{r['TRANSACTION TYPE']}_{r['COUNTRY']}_{i}"
        traded_energy = 0.0
        was_traded = 0

        if r["TRANSACTION TYPE"] == "SELL":
            if unit_name in n.generators.index:
                traded_energy = abs(n.generators_t.p.loc[0, unit_name])
                was_traded = int(traded_energy > 1e-6)
        elif r["TRANSACTION TYPE"] == "BUY":
            if unit_name in n.loads.index:
                traded_energy = n.loads_t.p.loc[0, unit_name]
                was_traded = int(traded_energy > 1e-6)

        trading_results.append({
            **h.to_dict(),
            "Bidding Area": "MI",
            "Agent": r.get("AGENT", ""),
            "Unit": r.get("UNIT", ""),
            "Country": r["COUNTRY"],
            "Technology": r["TECHNOLOGY"],
            "Capacity (MW)": r["BID ENERGY (MWH)"],
            "Transaction Type": r["TRANSACTION TYPE"],
            "Bid Price (EUR/MWh)": r["BID PRICE (EUR/MWH)"],
            "Bid Energy (MWh)": r["BID ENERGY (MWH)"],
            "Was Traded": was_traded,
            "Clearing Price (EUR/MWh)": price_pt if r["COUNTRY"] == "PT" else price_es,
            "Traded Energy (MWh)": traded_energy
        })

# =========================================================
# 5. EXPORT EXCEL
# =========================================================
session_df = pd.DataFrame(session_results)
trading_df = pd.DataFrame(trading_results)

with pd.ExcelWriter("SDAC_MIBEL_24H_RESULTS.xlsx", engine="xlsxwriter") as writer:
    session_df.to_excel(writer, sheet_name="Session Results", index=False)
    trading_df.to_excel(writer, sheet_name="Trading Results Detailed", index=False)

print("\n✔ 24h Day-Ahead MIBEL Simulation Completed Successfully!")

# =========================================================
# 6. PLOT 24H PRICE COMPARISON
# =========================================================
plt.figure(figsize=(15,6))
bar_width = 0.4
x = np.arange(len(session_df))

plt.bar(x - bar_width/2, session_df["Price_PT (EUR/MWh)"], width=bar_width, label="PT", color='skyblue')
plt.bar(x + bar_width/2, session_df["Price_ES (EUR/MWh)"], width=bar_width, label="ES", color='salmon')

plt.xlabel("Hour (PERIOD)")
plt.ylabel("Market Price (EUR/MWh)")
plt.title("24h Day-Ahead Market Prices: PT vs ES")
plt.xticks(x, session_df["PERIOD"], rotation=45)
plt.legend()
plt.grid(axis='y', linestyle='--', alpha=0.7)
plt.tight_layout()
plt.show()