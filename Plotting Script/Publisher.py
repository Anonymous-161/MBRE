import pandas as pd
import matplotlib.pyplot as plt

# Custom parameter settings
file_path = 'Data Table/Publisher.xlsx'
category_column = 'Publisher'
value_column = 'Number of papers'
color_palette = ['#3E87BA', '#5DA0C7', '#7EB9DE', '#A0D2F5']
font_size = 14
save_path = 'Publisher.png'

# Read data
df = pd.read_excel(file_path)
categories = df[category_column]
values = df[value_column]
total = values.sum()

# Create canvas
fig, ax = plt.subplots(figsize=(8, 5))

# Custom label function
def autopct_format(values):
    def inner_autopct(pct):
        val = int(round(pct * sum(values) / 100))
        return f'{val}\n({pct:.1f}%)'
    return inner_autopct

# Plot pie chart (correct key parameters)
wedges, texts, autotexts = ax.pie(
    values,
    labels=categories,  # Display category names
    colors=color_palette,
    startangle=90,
    autopct=autopct_format(values),
    pctdistance=0.75,
    wedgeprops={
        'edgecolor': 'white',
        'linewidth': 0.8
    },
    textprops={  # Uniformly set text style
        'fontsize': font_size-1,
        'fontweight': 'semibold',
        'color': '#2F2F2F'
    }
)

# Set percentage label style
plt.setp(autotexts,
        fontsize=font_size-1,
        color='white',
        fontweight='bold')

# Create legend and set style
legend = ax.legend(
    wedges,                      # Use pie chart wedge objects
    categories,                  # Display category names
    #title="Publisher",           # Legend title
    loc="center left",           # Position at center left
    bbox_to_anchor=(1, 0.5),     # Place legend at the middle of the right side of the canvas
    frameon=False,               # No border
    fontsize=font_size-2,        # Font size 2pt smaller than main labels
    title_fontsize=font_size,   # Title font size consistent with main labels
    labelspacing=1.2             # Increase label spacing
)

# Output chart
ax.axis('equal')
plt.tight_layout()
plt.savefig(save_path, dpi=300, bbox_inches='tight')
plt.show()