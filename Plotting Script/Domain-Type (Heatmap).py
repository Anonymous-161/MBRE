import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Parameter settings
excel_file = "Data Table/Domain-Type.xlsx"
sheet_name = "domain"
x_col = "Domain"
y_col = "Type"
output_img = "sorted_heatmap2.png"

# Read and process data
df = pd.read_excel(excel_file, sheet_name=sheet_name)
df_clean = df[[x_col, y_col]].dropna()

# Generate cross table and sort
cross_table = pd.crosstab(df_clean[x_col], df_clean[y_col])

# Calculate sorting indices
row_totals = cross_table.sum(axis=1).sort_values(ascending=False)  # Row totals in descending order
col_totals = cross_table.sum(axis=0).sort_values(ascending=True)    # Column totals in ascending order

# Reorder the cross table
sorted_cross = cross_table.reindex(
    index=row_totals.index,
    columns=col_totals.index
)

# Create canvas
fig, ax = plt.subplots(figsize=(12, 9))

# Plot heatmap
sns.heatmap(
    sorted_cross,
    annot=True,
    fmt="d",
    cmap="YlGnBu",
    linewidths=0.5,
    cbar=False,
    ax=ax
)
# Adjust left margin and label position
plt.subplots_adjust(left=0.1)  # Original value is usually between 0.1-0.2, decreasing this value will shift the whole to the left
# Adjust distance between y-axis labels and heatmap
ax.yaxis.set_tick_params(pad=20)  # Default 20, decreasing this value will shift the text to the right
# Add row totals on the left (same side as y-axis)
for y, (domain, total) in enumerate(row_totals.items()):
    ax.text(
        -0.3,  # Adjust X coordinate to the left
        y + 0.5,
        f"{int(total)}",
        ha='right',  # Right alignment
        va='center',
        fontsize=10,
        color='darkred'
    )

# Add column totals (keep original position)
col_totals = sorted_cross.sum(axis=0)
for x, (type_name, total) in enumerate(col_totals.items()):
    ax.text(
        x + 0.5,
        -0.5,
        f"{int(total)}",
        ha='center',
        va='center',
        fontsize=10,
        color='darkblue',
        rotation=45
    )

# Add color bar and format adjustments
cbar = ax.figure.colorbar(ax.collections[0], ax=ax, pad=0.01)
cbar.ax.set_ylabel('frequency of occurrence', rotation=270, labelpad=15)

ax.xaxis.set_label_position('top')
ax.xaxis.tick_top()
plt.xticks(rotation=45, ha='left')
plt.yticks(rotation=0)

# Adjust display range
ax.set_xlim(-0.8, sorted_cross.shape[1] + 1.2)  # Expand left space
ax.set_ylim(sorted_cross.shape[0] + 0.5, -1.5)
# Ensure x-axis does not squeeze left space

# Adjust Domain axis to move right

# Adjust margins (modify subplots_adjust parameters)
plt.subplots_adjust(
    left=0.25,  # Reduce left blank area <<< Key parameter
    right=0.85
)
ax.yaxis.set_tick_params(pad=15)  # Default 20, decreasing the value moves to the right
plt.tight_layout()
plt.savefig(output_img, dpi=300, bbox_inches='tight')
plt.show()