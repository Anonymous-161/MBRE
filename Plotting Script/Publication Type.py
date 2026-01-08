import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

# Custom parameter settings
file_path = 'Data Table/Publication Type.xlsx'  # Excel file path
category_column = 'Publication Type'  # Category column name in the data
value_column = 'Number of papers'  # Value column name in the data
bar_color = '#3E87BA'  # Bar chart color (light navy blue)
bar_width = 0.30  # Adjust bar width (slightly thinner)
grid_color = 'lightgray'  # Grid line color
grid_linewidth = 0.5  # Grid line width
grid_linestyle = '--'  # Grid line style
font_size_title = 14  # Title font size
font_size_labels = 14  # Label font size
save_path = 'Publication Type.png'  # Path to save the image

# Read Excel data
df = pd.read_excel(file_path)

# Get data
categories = df[category_column]
values = df[value_column]

# Create chart
fig, ax = plt.subplots(figsize=(8, 5))

# Plot bar chart
bars = ax.bar(categories, values, color=bar_color, width=bar_width)

# Set grid background
ax.set_facecolor('white')  # Set background color to white
ax.grid(True, which='both', axis='both', linestyle=grid_linestyle, linewidth=grid_linewidth, color=grid_color)

# Set x-axis to display each value with an interval of 0.5 between coordinates
# start = np.min(categories) - 0.5  # Start point
# end = np.max(categories) + 0.5    # End point
# ax.set_xticks(np.arange(start, end, 0.5))  # Set x-axis ticks with 0.5 interval

# Set y-axis range from 0 to 50
ax.set_ylim(0, 60)

# Set title and labels
# ax.set_title('Bar Chart from Excel Data', fontsize=font_size_title)
# ax.set_xlabel('Quality score', fontsize=font_size_labels)
# ax.set_ylabel('Number of papers', fontsize=font_size_labels)

ax.tick_params(
    axis='x',                # Only operate on x-axis (y-axis is hidden)
    which='major',           # Major ticks
    labelsize=12,            # Font size
    colors='#2F2F2F',       # Font color (dark gray)
    pad=6                    # Spacing between labels and axis
)
# Set x-axis tick labels to bold individually
for label in ax.get_xticklabels():
    label.set_fontweight('bold')  # Bold style: 'bold'/'semibold'

ax.tick_params(left=False, labelleft=False)
# Display x-axis and y-axis spines
ax.spines['bottom'].set_visible(True)
ax.spines['left'].set_visible(False)
ax.spines['right'].set_visible(False)  # Remove right spine
ax.spines['top'].set_visible(False)  # Remove top spine

# Add data labels above each bar
ax.bar_label(bars, fontsize=14, padding=3, color='black')

# Save the image
plt.tight_layout()
plt.savefig(save_path, dpi=300)  # Save as PNG file with 300 DPI high resolution

# Display the chart
plt.show()