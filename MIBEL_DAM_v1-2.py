import pandas as pd
import os
import pypsa
import numpy as np
import matplotlib.pyplot as plt

np.random.seed(42)

# =========================================================
# 1. LOAD DATA
# =========================================================
file_path = "Input_Exemplo.xlsx"
base_name, ext = os.path.splitext(file_path)
output_file = f"{base_name}_MarketResults{ext}"
df = pd.read_excel(file_path)

df.columns = df.columns.str.strip().str.upper()
for c in ["TRANSACTION TYPE", "COUNTRY", "TECHNOLOGY", "UNIT", "AGENT"]:
    if c in df.columns:
        df[c] = df[c].astype(str).str.strip().str.upper()

# Alias price + tie-breaker
df["BID PRICE (EUR/MWH)"] = df["BID PRICE RANDOM LNEG 2 (EUR/MWH)"]
df["BID PRICE (EUR/MWH)"] += 0.001 * np.random.rand(len(df))

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

# =========================================================
# 3. RESULT CONTAINERS
# =========================================================
session_results = []
trading_results = []

MANDATORY_PRICE = 3880
EPS = 1e-3

# =========================================================
# 4. HOURLY MARKET CLEARING
# =========================================================
for _, h in hours.iterrows():

    period = h["PERIOD"]

    print(
        f"\n===== SIMULATION HOUR {period} =====\n"
        f"PERIOD OF YEAR: {h['PERIOD OF YEAR']}\n"
        f"Date: {h['YEAR']}-{h['MONTH']:02d}-{h['DAY']:02d}\n"
        f"SESSION: {h['SESSION']} | PERIOD: {period}\n"
    )

    df_h = df[df["PERIOD"] == period].copy()

    pt_sell = df_h[(df_h.COUNTRY == "PT") & (df_h["TRANSACTION TYPE"] == "SELL")]
    es_sell = df_h[(df_h.COUNTRY == "ES") & (df_h["TRANSACTION TYPE"] == "SELL")]
    pt_buy  = df_h[(df_h.COUNTRY == "PT") & (df_h["TRANSACTION TYPE"] == "BUY")]
    es_buy  = df_h[(df_h.COUNTRY == "ES") & (df_h["TRANSACTION TYPE"] == "BUY")]

    # -------------------------------
    # Network
    # -------------------------------
    n = pypsa.Network()
    n.set_snapshots([0])
    n.add("Carrier", "electricity")
    n.add("Bus", "PT", carrier="electricity")
    n.add("Bus", "ES", carrier="electricity")

    intercap = float(df_h.get("INTERCONNECTION", pd.Series([3800])).iloc[0])
    n.add(
        "Link",
        "PT_ES",
        bus0="PT",
        bus1="ES",
        p_nom=intercap,
        p_min_pu=-1,
        efficiency=1.0
    )

    # -------------------------------
    # SELL bids
    # -------------------------------
    for i, r in pd.concat([pt_sell, es_sell]).iterrows():
        e = float(r["BID ENERGY (MWH)"])
        if e <= 0:
            continue
        n.add(
            "Generator",
            f"SELL_{r['COUNTRY']}_{i}",
            bus=r["COUNTRY"],
            p_nom=e,
            marginal_cost=float(r["BID PRICE (EUR/MWH)"])
        )

    # -------------------------------
    # BUY bids
    # -------------------------------
    for i, r in pd.concat([pt_buy, es_buy]).iterrows():
        e = float(r["BID ENERGY (MWH)"])
        p = float(r["BID PRICE (EUR/MWH)"])
        if e <= 0:
            continue
        if p >= MANDATORY_PRICE:
            n.add("Load", f"BUY_{r['COUNTRY']}_{i}", bus=r["COUNTRY"], p_set=e)
        else:
            n.add(
                "Generator",
                f"FLEX_{r['COUNTRY']}_{i}",
                bus=r["COUNTRY"],
                p_nom=e,
                p_min_pu=-1,
                p_max_pu=0,
                marginal_cost=p
            )

    # -------------------------------
    # Solve
    # -------------------------------
    n.optimize(solver_name="glpk")

    price_pt = n.buses_t.marginal_price.loc[0, "PT"]
    price_es = n.buses_t.marginal_price.loc[0, "ES"]

    link_flow = n.links_t.p0.loc[0, "PT_ES"]
    congested = abs(abs(link_flow) - intercap) < EPS

    print(
        f">>> Prices | PT: {price_pt:.4f} EUR/MWh | ES: {price_es:.4f} EUR/MWh | "
        f"Congested: {'YES' if congested else 'NO'}"
    )
    print(
        f">>> Interconnection | Capacity: {intercap:.0f} MW | "
        f"PT→ES: {max(link_flow,0):.2f} MW | ES→PT: {max(-link_flow,0):.2f} MW"
    )

    # -------------------------------
    # Quantities
    # -------------------------------
    sell_mask = n.generators.index.str.startswith("SELL")
    flex_mask = n.generators.index.str.startswith("FLEX")

    total_supply = n.generators_t.p.loc[0, sell_mask].sum()
    total_flex_demand = -n.generators_t.p.loc[0, flex_mask].sum()
    total_mandatory_demand = n.loads_t.p.loc[0].sum()
    total_demand = total_flex_demand + total_mandatory_demand

    # -------------------------------
    # Welfare
    # -------------------------------
    producer_surplus = 0.0
    consumer_surplus = 0.0

    for g in n.generators.index:
        q = n.generators_t.p.loc[0, g]
        if abs(q) < 1e-9:
            continue
        bid = n.generators.loc[g, "marginal_cost"]
        zone = n.generators.loc[g, "bus"]
        market_price = price_pt if zone == "PT" else price_es

        if g.startswith("SELL") and q > 0:
            producer_surplus += (market_price - bid) * q
        if g.startswith("FLEX") and q < 0:
            consumer_surplus += (bid - market_price) * (-q)

    total_welfare = producer_surplus + consumer_surplus
    congestion_rent = abs(link_flow) * abs(price_pt - price_es)

    # -------------------------------
    # Session results
    # -------------------------------
    session_results.append({
        **h.to_dict(),
        "Bidding Area": "MI",
        "Price_PT (EUR/MWh)": price_pt,
        "Price_ES (EUR/MWh)": price_es,
        "Congested": int(congested),
        "PT→ES Flow (MW)": max(link_flow, 0),
        "ES→PT Flow (MW)": max(-link_flow, 0),
        "Total Supply (MWh)": total_supply,
        "Total Demand (MWh)": total_demand,
        "Producer Surplus (€)": producer_surplus,
        "Consumer Surplus (€)": consumer_surplus,
        "Total Welfare (€)": total_welfare,
        "Congestion Rent (€)": congestion_rent
    })

    # -------------------------------
    # Trading Results Detailed
    # -------------------------------
    for i, r in df_h.iterrows():

        traded_energy = 0.0
        was_traded = 0

        if r["TRANSACTION TYPE"] == "SELL":
            name = f"SELL_{r['COUNTRY']}_{i}"
            if name in n.generators.index:
                traded_energy = n.generators_t.p.loc[0, name]
                was_traded = int(traded_energy > 1e-6)

        elif r["TRANSACTION TYPE"] == "BUY":
            if r["BID PRICE (EUR/MWH)"] >= MANDATORY_PRICE:
                name = f"BUY_{r['COUNTRY']}_{i}"
                if name in n.loads.index:
                    traded_energy = n.loads_t.p.loc[0, name]
                    was_traded = int(traded_energy > 1e-6)
            else:
                name = f"FLEX_{r['COUNTRY']}_{i}"
                if name in n.generators.index:
                    traded_energy = -n.generators_t.p.loc[0, name]
                    was_traded = int(traded_energy > 1e-6)

        clearing_price_unit = price_pt if r["COUNTRY"] == "PT" else price_es

        trading_results.append({
            "PERIOD OF YEAR": h["PERIOD OF YEAR"],
            "YEAR": h["YEAR"],
            "MONTH": h["MONTH"],
            "DAY": h["DAY"],
            "SESSION": h["SESSION"],
            "PERIOD": h["PERIOD"],
            "Bidding Area": "MI",
            "Agent": r.get("AGENT", ""),
            "Unit": r.get("UNIT", ""),
            "Country": r["COUNTRY"],
            "Technology": r.get("TECHNOLOGY", ""),
            "Capacity (MW)": r["BID ENERGY (MWH)"],
            "Transaction Type": r["TRANSACTION TYPE"],
            "Bid Price (EUR/MWh)": r["BID PRICE (EUR/MWH)"],
            "Bid Energy (MWh)": r["BID ENERGY (MWH)"],
            "Was Traded": was_traded,
            "Clearing Price (EUR/MWh)": clearing_price_unit,
            "Traded Energy (MWh)": traded_energy
        })

# =========================================================
# 5. EXPORT RESULTS
# =========================================================

session_df = pd.DataFrame(session_results)
trading_df = pd.DataFrame(trading_results)

with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
    session_df.to_excel(writer, sheet_name="Session Results", index=False)
    trading_df.to_excel(writer, sheet_name="Trading Results Detailed", index=False)

print("\n✔ 24h Day-Ahead MIBEL Simulation Completed Successfully!")

# =========================================================
# 6. PRICE PLOT
# =========================================================
plt.figure(figsize=(15,6))
x = np.arange(len(session_df))
plt.bar(x - 0.2, session_df["Price_PT (EUR/MWh)"], width=0.4, label="PT")
plt.bar(x + 0.2, session_df["Price_ES (EUR/MWh)"], width=0.4, label="ES")
plt.xticks(x, session_df["PERIOD"], rotation=45)
plt.ylabel("EUR/MWh")
plt.title("24h Day-Ahead Market Prices: PT vs ES")
plt.legend()
plt.grid(axis="y", linestyle="--", alpha=0.6)
plt.tight_layout()
plt.show()