import matplotlib.pyplot as plt
import pandas as pd

# Read Excel data
excel_path = "Data Table/Technology Score.xlsx"
df = pd.read_excel(excel_path)

# Column name validation and cleaning
required_columns = ['Technologies','Learnability','Expressiveness','Collaboration','Toolchain','Scalability','Cost & Resources']
if not set(required_columns).issubset(df.columns):
    df.columns = df.columns.str.replace(r'[ &]', '', regex=True)
    required_columns_clean = [col.replace(' ', '').replace('&','') for col in required_columns]
    if set(required_columns_clean).issubset(df.columns):
        df = df.rename(columns=dict(zip(required_columns_clean, required_columns)))
    else:
        missing = set(required_columns) - set(df.columns)
        raise ValueError(f"Missing required columns: {missing}")

# Standardize data format
dimensions = required_columns[1:]
df[dimensions] = df[dimensions].apply(pd.to_numeric, errors='coerce').fillna(0).astype(int)

# Calculate total score and sort by total score in descending order
df['Total'] = df[dimensions].sum(axis=1)
df = df.sort_values(by='Total', ascending=True).reset_index(drop=True)

# Coordinate mapping
tech_map = {tech:i for i,tech in enumerate(df['Technologies'])}
dim_map = {dim:i for i,dim in enumerate(dimensions)}

# Create plotting data
plot_data = []
for _, row in df.iterrows():
    for dim in dimensions:
        plot_data.append({
            'x': dim_map[dim],
            'y': tech_map[row['Technologies']],
            'size': row[dim],
            'score': row[dim],
            'tech': row['Technologies'],
            'total': row['Total']
        })
df_plot = pd.DataFrame(plot_data)

# Chart parameter settings
plt.figure(figsize=(10, 7.5))
ax = plt.gca()
BUBBLE_BASE_SIZE = 350
TOTAL_SCORE_X = -0.60

# Plot bubble chart
sc = ax.scatter(
    x='x',
    y='y',
    s=df_plot['size']*BUBBLE_BASE_SIZE,
    data=df_plot,
    edgecolor='black',
    linewidth=0.8,
    alpha=0.85,
    zorder=2
)

# Add score labels
for _, row in df_plot.iterrows():
    ax.text(
        row['x'], row['y'],
        str(row['score']),
        ha='center', va='center',
        fontsize=10,
        fontweight='bold',
        color='white',
        zorder=3
    )

# Add total score column
for tech, idx in tech_map.items():
    total = df[df['Technologies'] == tech]['Total'].values[0]
    ax.text(
        TOTAL_SCORE_X, idx,
        f"{total}",
        ha='right', va='center',
        fontsize=10.5,
        fontweight='bold',
        color='#1f77b4'
    )

# Calculate total score for each dimension
dim_totals = df[dimensions].sum()

# Display total score above each dimension
for i, dim in enumerate(dimensions):
    ax.text(
        i, len(tech_map) + 0.06,  # Position: above dimension
        f"{dim_totals[dim]}",
        ha='center', va='bottom',
        fontsize=10,
        fontweight='bold',
        color='#1f77b4'
    )

# Axis settings
ax.set_xticks(range(len(dimensions)))
ax.set_xticklabels(
    dimensions,
    rotation=40,
    ha='right',
    rotation_mode='anchor',
    fontweight='bold',
    fontsize=11
)
ax.set_yticks(range(len(tech_map)))
ax.set_yticklabels(
    df['Technologies'],
    fontweight='bold',
    fontsize=11
)

# Layout optimization
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.set_xlim(-0.9, len(dimensions)-0.4)
ax.set_ylim(-0.5-0.2, len(tech_map)-0.5+0.2)
plt.grid(True, linestyle=':', alpha=0.5, zorder=1)

# Final layout adjustment
plt.tight_layout()
plt.subplots_adjust(
    left=0.13,
    right=0.97,
    top=0.95,
    bottom=0.18
)

# Save chart
plt.savefig("technologies_bubble_chart2.png", dpi=300, bbox_inches='tight')
plt.show()