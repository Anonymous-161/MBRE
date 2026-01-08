import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import numpy as np

# ================= Configuration Area =================
file_path = 'Data Table/Region_new.xlsx'
sheet_name = 0
category_column = 'Region'
score_column = 'Score'
count_column = 'Number of papers'

# Color Settings
color_score = '#3E87BA'    # Dark blue (Left Axis - Score)
color_count = '#9AC9DB'    # Light blue (Right Axis - Paper Count)

bar_width = 0.35           # Width of single bar
font_size_labels = 12      # Axis label font size
save_path = 'Region_Dual_Axis_Final.png'
# ===========================================

# 1. Read Data
try:
    df = pd.read_excel(file_path, sheet_name=sheet_name)
except Exception as e:
    print(f"Failed to read file: {e}")
    exit()

categories = df[category_column]
scores = df[score_column]
counts = df[count_column]
x = np.arange(len(categories))

# 2. Create Canvas
fig, ax1 = plt.subplots(figsize=(12, 6))
ax2 = ax1.twinx()

# ================= Core Modification: Adjust Layer Order (Z-order Sandwich) =================
# Logic: Place ax1 (left axis) on top of ax2 (right axis) to prevent ax2 elements from obscuring ax1 labels
# However, grid lines should be drawn on ax2 (bottom layer), otherwise grid lines will block ax2 bars
ax1.set_zorder(10)          # Place ax1 on the top layer
ax1.patch.set_visible(False)# Make ax1 background transparent to avoid blocking ax2 below
ax2.set_zorder(5)           # Place ax2 on the bottom layer
# =======================================================================

# 3. Plot Bar Chart
# rects1 (Score): Drawn on ax1 (top layer)
rects1 = ax1.bar(x - bar_width/2, scores, width=bar_width, label='Fractional Score', color=color_score)
# rects2 (Count): Drawn on ax2 (bottom layer)
rects2 = ax2.bar(x + bar_width/2, counts, width=bar_width, label='Number of papers', color=color_count)

# 4. Set Grid Lines (Drawn on bottom layer ax2)
ax1.grid(True, which='both', axis='both', linestyle='--', linewidth=0.5, color='lightgray', zorder='15')
# Key: Ensure grid lines are behind the bars
# ax1.set_axisbelow(True)

# 5. Set Left Axis (Score)
ax1.set_xlabel('Region', fontsize=font_size_labels)
ax1.set_ylabel('Fractional Score', fontsize=font_size_labels, color=color_score, fontweight='bold')
ax1.tick_params(axis='y', labelcolor=color_score)
ax1.set_ylim(0, 20)

# 6. Set Right Axis (Paper Count)
ax2.set_ylabel('Number of papers', fontsize=font_size_labels, color=color_count, fontweight='bold')
ax2.tick_params(axis='y', labelcolor=color_count)
ax2.yaxis.set_major_locator(ticker.MaxNLocator(integer=True))
ax2.set_ylim(0, 20)

# 7. Set Common X-axis
ax1.set_xticks(x)
ax1.set_xticklabels(categories, rotation=50, ha='right', fontsize=11)

# 8. Merge Legends
lines1, labels1 = ax1.get_legend_handles_labels()
lines2, labels2 = ax2.get_legend_handles_labels()
# Place legend at top-left or top-right to avoid obscuring data
ax1.legend(lines1 + lines2, labels1 + labels2, frameon=False, fontsize=11, loc='upper right')

# 9. Remove Top Border
ax1.spines['top'].set_visible(False)
ax2.spines['top'].set_visible(False)

# 10. Add Value Labels (This is the most critical step)
# Since ax1 is on the top layer, its labels are naturally on top
ax1.bar_label(rects1, padding=2, fontsize=10, color=color_score, fmt='%.1f', fontweight='bold')

# ax2 is on the bottom layer, but we can manually increase the zorder of its labels, or rely on ax1's transparent background
# The trick here is to draw directly, as the bars are staggered and won't obscure each other as long as the background is transparent
ax2.bar_label(rects2, padding=2, fontsize=10, color=color_count, fmt='%.0f', fontweight='bold')

# 11. Save and Display
plt.tight_layout()
plt.savefig(save_path, dpi=300)
print(f"Image saved as: {save_path}")
plt.show()