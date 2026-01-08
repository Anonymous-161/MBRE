import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker

# Custom parameter settings
file_path = 'Data Table/Region.xlsx'  # Path to the uploaded file
category_column = 'Region'  # Category column name in your data
value_column = 'Number of papers'  # Value column name in your data
bar_color = '#3E87BA'  # Bar chart color (light navy blue)
bar_width = 0.5  # Adjust bar width (slightly thinner)
grid_color = 'lightgray'  # Grid line color
grid_linewidth = 0.5  # Grid line width
grid_linestyle = '--'  # Grid line style
font_size_title = 14  # Title font size
font_size_labels = 12  # Label font size
save_path = 'Region.png'  # Path to save the image

# Read Excel data
df = pd.read_excel(file_path)

# Get data
categories = df[category_column]
values = df[value_column]

# Create chart
fig, ax = plt.subplots(figsize=(11, 5))

# Plot bar chart
bars = ax.bar(categories, values, color=bar_color, width=bar_width)

# Set grid background
ax.set_facecolor('white')  # Set background color to white
ax.grid(True, which='both', axis='both', linestyle=grid_linestyle, linewidth=grid_linewidth, color=grid_color)

# Set x-axis to display region names
ax.set_xticks(categories)  # Set x-axis ticks to region names

# Set y-axis range to adjust automatically based on data
ax.set_ylim(0, max(values) + 3)  # Set y-axis maximum to data maximum + 5 to prevent bars from touching the top

# Set y-axis to display integers
ax.yaxis.set_major_locator(ticker.MaxNLocator(integer=True))

# Set title and labels
ax.set_xlabel('Region', fontsize=font_size_labels)
ax.set_ylabel('Number of papers', fontsize=font_size_labels)

# Display x-axis and y-axis spines
ax.spines['bottom'].set_visible(True)
ax.spines['left'].set_visible(True)
ax.spines['right'].set_visible(False)  # Remove right spine
ax.spines['top'].set_visible(False)  # Remove top spine

# Add data labels above each bar
ax.bar_label(bars, fontsize=12, padding=3, color='black')

# Rotate x-axis tick labels
plt.xticks(rotation=50,ha='right')

# Save the image
plt.tight_layout()
plt.savefig(save_path, dpi=300)  # Save as PNG file with 300 DPI high resolution

# Display the chart
plt.show()