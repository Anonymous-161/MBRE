import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

# Custom parameter settings
file_path = 'Data Table/Score Distribution.xlsx'  # Excel file path
category_column = 'Score'  # Category column name
value_column = 'Number of papers'  # Value column name
bar_color = '#3E87BA'  # Bar chart color (light navy blue)
bar_width = 0.25  # Bar width (slightly narrower)
grid_color = 'lightgray'  # Grid line color
grid_linewidth = 0.5  # Grid line width
grid_linestyle = '--'  # Grid line style
font_size_title = 14  # Title font size
font_size_labels = 12  # Label font size
save_path = 'quality_scores.png'  # Image save path

# Read Excel data
df = pd.read_excel(file_path)

# Get data
categories = df[category_column]
values = df[value_column]

# Create chart
fig, ax = plt.subplots(figsize=(8, 6))

# Plot bar chart
bars = ax.bar(categories, values, color=bar_color, width=bar_width)

# Set grid background
ax.set_facecolor('white')  # Set background to white
ax.grid(True, which='both', axis='both', linestyle=grid_linestyle, linewidth=grid_linewidth, color=grid_color)

# Set x-axis to display each value with 0.5 interval between ticks
start = np.min(categories) - 0.5  # Start point
end = np.max(categories) + 0.5    # End point
ax.set_xticks(np.arange(start, end, 0.5))  # Set x-axis ticks with 0.5 interval

# Set y-axis range from 0 to 50
ax.set_ylim(0, 50)

# Set title and labels
# ax.set_title('Bar Chart from Excel Data', fontsize=font_size_title)
ax.set_xlabel('Quality score', fontsize=font_size_labels)
ax.set_ylabel('Number of papers', fontsize=font_size_labels)

# Show x and y axis spines
ax.spines['bottom'].set_visible(True)
ax.spines['left'].set_visible(True)
ax.spines['right'].set_visible(False)  # Remove right spine
ax.spines['top'].set_visible(False)  # Remove top spine

# Add value labels above each bar
ax.bar_label(bars, fontsize=12, padding=3, color='black')

# Save image (300 DPI high resolution)
plt.tight_layout()
plt.savefig(save_path, dpi=300)

# Display chart
plt.show()