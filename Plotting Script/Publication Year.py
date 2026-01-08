import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker

# Custom parameter settings
file_path = 'Data Table/Publication Year.xlsx'  # Path to the uploaded file
category_column = 'Publication Year'  # Category column name in the data (e.g., year)
value_column = 'Number of papers'  # Value column name in the data
line_color = '#3E87BA'  # Line chart color (light navy blue)
grid_color = 'lightgray'  # Grid line color
grid_linewidth = 1  # Grid line width
grid_linestyle = '--'  # Grid line style
font_size_title = 14  # Title font size
font_size_labels = 12  # Label font size
save_path = 'line_chart.png'  # Path to save the image

# Read Excel data
df = pd.read_excel(file_path)

# Get data
categories = df[category_column]
values = df[value_column]

# Create chart
fig, ax = plt.subplots(figsize=(13, 5))  # Increase chart width

# Plot line chart
ax.plot(categories, values, marker='o', color=line_color, linestyle='-', linewidth=3)

# Set grid background
ax.set_facecolor('white')  # Set background color to white
ax.grid(True, which='both', axis='both', linestyle=grid_linestyle, linewidth=grid_linewidth, color=grid_color)

# Set x-axis to display years
ax.set_xticks(categories)  # Set x-axis ticks to years

# Set y-axis range to adjust automatically based on data
ax.set_ylim(0, max(values) + 3)  # Set y-axis maximum to data maximum + 3 to avoid overcrowded plot

# Set y-axis to display integers
ax.yaxis.set_major_locator(ticker.MaxNLocator(integer=True))

# Set title and labels
ax.set_xlabel('Publication Year', fontsize=font_size_labels)
ax.set_ylabel('Number of papers', fontsize=font_size_labels)
# ax.set_title('Number of Papers per Year', fontsize=font_size_title)

# Display x-axis and y-axis spines
ax.spines['bottom'].set_visible(True)
ax.spines['left'].set_visible(True)
ax.spines['right'].set_visible(False)  # Remove right spine
ax.spines['top'].set_visible(False)  # Remove top spine

# Add value labels above each data point
for i, value in enumerate(values):
    ax.text(categories[i], value + 0.3, str(value), ha='center', fontsize=12)

# Rotate x-axis tick labels
plt.xticks(rotation=40, ha='right')

# Save the image
plt.tight_layout()
plt.savefig(save_path, dpi=300)  # Save as PNG file with 300 DPI high resolution

# Display the chart
plt.show()