# ===================== IMPORT REQUIRED LIBRARIES =====================
# matplotlib → for bar charts
# pandas → for Excel file reading
# FPDF → for PDF generation
# os → to delete temporary image files

import matplotlib.pyplot as plt
import pandas as pd
from fpdf import FPDF
import os


# ===================== DEFINE 16 MAHAVASTU ZONES =====================
# Order is IMPORTANT (used everywhere: charts, tables, Excel mapping)

zones = [
    "N","NNE","NE","ENE",
    "E","ESE","SE","SSE",
    "S","SSW","SW","WSW",
    "W","WNW","NW","NNW"
]


# ===================== PRAKRITI ZONE MAPPING =====================
# Which zones belong to Fire / Water / Air

prakriti_map = {
    "Fire": ["E","ESE","SE","SSE","S","SSW"],
    "Water": ["NNW","N","NNE","NE","ENE"],
    "Air": ["SW","WSW","W","WNW","NW"]
}

# Light pastel colors for Prakriti bar chart
prakriti_colors = {
    "Fire": "#ffcccb",     # light red
    "Water": "#cce5ff",    # light blue
    "Air": "#d5f5e3"       # light green
}


# ===================== USER INPUT METHOD =====================
print("Choose input method:")
print("1. Manual Entry")
print("2. Load from Excel")

choice = input("Enter choice (1 or 2): ")

# Dictionary to store zone → area values
zone_data = {}


# ===================== MANUAL INPUT =====================
if choice == "1":
    print("\nEnter area for each zone:")
    for z in zones:
        zone_data[z] = float(input(f"{z}: "))


# ===================== EXCEL INPUT =====================
elif choice == "2":
    file_path = input("Enter Excel file path: ")

    # Excel must have columns: Zone | Area
    df = pd.read_excel(file_path)

    # Reindex ensures correct zone order
    df = df.set_index("Zone").reindex(zones)

    # Convert to dictionary
    zone_data = df["Area"].to_dict()

else:
    print("Invalid choice. Exiting.")
    exit()


# Convert zone values to list (used for calculations)
areas = list(zone_data.values())


# ===================== BALANCE LINE CALCULATIONS =====================
# LOB  → Line of Balance (Average)
# ULOB → Upper Line of Balance
# LLOB → Lower Line of Balance

LOB = sum(areas) / len(areas)
ULOB = (LOB + max(areas)) / 2
LLOB = (LOB + min(areas)) / 2


# ===================== AUTO REMARKS FOR EACH ZONE =====================
# High     → Above ULOB
# Balanced → Between LLOB & ULOB
# Low      → Below LLOB

remarks = {}

for z, v in zone_data.items():
    if v > ULOB:
        remarks[z] = "High"
    elif v < LLOB:
        remarks[z] = "Low"
    else:
        remarks[z] = "Balanced"


# ===================== PRAKRITI CALCULATION =====================
# Sum zone values element-wise

prakriti_strength = {}

for p, zlist in prakriti_map.items():
    prakriti_strength[p] = sum(zone_data[z] for z in zlist)

# Total of Fire + Water + Air
total_prakriti = sum(prakriti_strength.values())

# Convert Prakriti values to percentage
prakriti_percentage = {
    p: (v / total_prakriti) * 100
    for p, v in prakriti_strength.items()
}

# Sort Prakriti by dominance
prakriti_sorted = sorted(
    prakriti_strength.items(),
    key=lambda x: x[1],
    reverse=True
)

# Final Prakriti name (e.g. Fire-Water-Air)
final_prakriti = "-".join(p[0] for p in prakriti_sorted)


# ===================== BAR CHART 1: ZONAL STRENGTH =====================
plt.figure(figsize=(12,6))

bars = plt.bar(
    zones,
    areas,
    color="#aed6f1",
    edgecolor="black"
)

# Display value above each bar
for bar, val in zip(bars, areas):
    plt.text(
        bar.get_x() + bar.get_width()/2,
        bar.get_height(),
        f"{val:.0f}",
        ha="center",
        va="bottom",
        fontsize=8
    )

# Draw balance lines
plt.axhline(LOB, color="blue", linestyle="--", linewidth=2,
            label=f"LOB = {LOB:.1f}")

plt.axhline(ULOB, color="red", linestyle="--", linewidth=2,
            label=f"ULOB = {ULOB:.1f}")

plt.axhline(LLOB, color="orange", linestyle="--", linewidth=2,
            label=f"LLOB = {LLOB:.1f}")

plt.title("MahaVastu Zonal Strength Analysis")
plt.xlabel("16 Zones")
plt.ylabel("Area / Strength")
plt.legend()

zone_chart = "zone_chart.png"
plt.savefig(zone_chart, dpi=300, bbox_inches="tight")
plt.close()


# ===================== BAR CHART 2: PRAKRITI =====================
plt.figure(figsize=(7,5))

bars = plt.bar(
    prakriti_strength.keys(),
    prakriti_strength.values(),
    color=[prakriti_colors[k] for k in prakriti_strength],
    edgecolor="black"
)

# Display values above bars
for bar, val in zip(bars, prakriti_strength.values()):
    plt.text(
        bar.get_x() + bar.get_width()/2,
        bar.get_height(),
        f"{val:.0f}",
        ha="center",
        va="bottom",
        fontsize=11,
        fontweight="bold"
    )

plt.title("Mahavastu Building Prakriti Bar Chart")
plt.ylabel("Total Zonal Strength")

prakriti_chart = "prakriti_chart.png"
plt.savefig(prakriti_chart, dpi=300, bbox_inches="tight")
plt.close()


# ===================== PDF REPORT GENERATION =====================
pdf = FPDF()
pdf.set_auto_page_break(auto=True, margin=15)


# ---------- PAGE 1: ZONAL REPORT ----------
pdf.add_page()
pdf.set_font("Arial", 'B', 16)
pdf.cell(0, 10, "MahaVastu Zonal Strength Report", ln=True, align="C")

pdf.ln(5)
pdf.image(zone_chart, x=10, w=190)

# Balance line full forms
pdf.ln(6)
pdf.set_font("Arial", size=12)
pdf.cell(0, 8, f"LOB (Line of Balance): {LOB:.2f}", ln=True)
pdf.cell(0, 8, f"ULOB (Upper Line of Balance): {ULOB:.2f}", ln=True)
pdf.cell(0, 8, f"LLOB (Lower Line of Balance): {LLOB:.2f}", ln=True)

# Zone table
pdf.ln(6)
pdf.set_font("Arial", 'B', 11)
pdf.cell(35, 8, "Zone", 1)
pdf.cell(45, 8, "Area", 1)
pdf.cell(110, 8, "Remark", 1, ln=True)

pdf.set_font("Arial", size=11)
for z in zones:
    pdf.cell(35, 8, z, 1)
    pdf.cell(45, 8, f"{zone_data[z]:.0f}", 1)
    pdf.cell(110, 8, remarks[z], 1, ln=True)


# ---------- PAGE 2: PRAKRITI REPORT ----------
pdf.add_page()
pdf.set_font("Arial", 'B', 16)
pdf.cell(0, 10, "Mahavastu Prakriti Analysis", ln=True, align="C")

pdf.ln(5)
pdf.image(prakriti_chart, x=30, w=150)

# Prakriti table
pdf.ln(8)
pdf.set_font("Arial", 'B', 12)
pdf.cell(50, 8, "Element", 1)
pdf.cell(50, 8, "Strength", 1)
pdf.cell(40, 8, "Percentage", 1, ln=True)

pdf.set_font("Arial", size=12)
for p in prakriti_strength:
    pdf.cell(50, 8, p, 1)
    pdf.cell(50, 8, f"{prakriti_strength[p]:.0f}", 1)
    pdf.cell(40, 8, f"{prakriti_percentage[p]:.1f}%", 1, ln=True)

# Final Prakriti
pdf.ln(6)
pdf.set_font("Arial", 'B', 12)
pdf.cell(0, 10, f"Final Building Prakriti: {final_prakriti}", ln=True)


# ===================== SAVE PDF WITH CUSTOM NAME =====================
custom_name = input("\nEnter PDF file name (without .pdf): ").strip()
if not custom_name:
    custom_name = "Mahavastu_Report"

pdf.output(f"{custom_name}.pdf")

# Remove temporary chart images
os.remove(zone_chart)
os.remove(prakriti_chart)

print(f"✅ {custom_name}.pdf generated successfully")
