import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker

# Custom parameter settings
file_path = 'Data Table/Score Details.xlsx'  # Excel file path
category_column = 'Dimension'  # Dimension column name
low_column = 'Not'  # Low score column name
medium_column = 'To some extend'  # Medium score column name
high_column = 'Yes'  # High score column name
bar_width = 0.5  # Bar width (slightly narrower)
grid_color = 'lightgray'  # Grid line color
grid_linewidth = 0.5  # Grid line width
grid_linestyle = '--'  # Grid line style
font_size_title = 14  # Title font size
font_size_labels = 12  # Label font size
save_path = 'QC_Score.png'  # Image save path

# Read Excel data
df = pd.read_excel(file_path)

# Get data
categories = df[category_column]
low_scores = df[low_column]
medium_scores = df[medium_column]
high_scores = df[high_column]

# Create chart (increase width)
fig, ax = plt.subplots(figsize=(10, 6))

# Color palette settings
low_color = '#ADC6E5'  # Light gray (low score)
medium_color = '#5FB1ED'  # Light blue (medium score)
high_color = '#3E87BA'  # Light navy blue (high score)

# Plot stacked horizontal bar chart
ax.barh(categories, low_scores, label='Not', color=low_color, height=bar_width)
ax.barh(categories, medium_scores, left=low_scores, label='To some extend', color=medium_color, height=bar_width)
ax.barh(categories, high_scores, left=low_scores + medium_scores, label='Yes', color=high_color, height=bar_width)

# Set grid background
ax.set_facecolor('white')  # Set background to white
ax.grid(True, which='both', axis='both', linestyle=grid_linestyle, linewidth=grid_linewidth, color=grid_color)

# Set y-axis range
ax.set_ylim(-0.5, len(categories) - 0.5)

# Set x-axis to integer values
ax.xaxis.set_major_locator(ticker.MaxNLocator(integer=True))

# Set title
# ax.set_title('Scores Distribution per Dimension', fontsize=font_size_title)

# Show x and y axis spines
ax.spines['bottom'].set_visible(True)
ax.spines['left'].set_visible(True)
ax.spines['right'].set_visible(False)  # Remove right spine
ax.spines['top'].set_visible(False)  # Remove top spine

# Add non-zero value labels in the middle of each stacked bar segment
for i, (low, medium, high) in enumerate(zip(low_scores, medium_scores, high_scores)):
    if low > 0:
        ax.text(low / 2, i, str(low), ha='center', va='center', fontsize=12, color='white')
    if medium > 0:
        ax.text(low + medium / 2, i, str(medium), ha='center', va='center', fontsize=12, color='white')
    if high > 0:
        ax.text(low + medium + high / 2, i, str(high), ha='center', va='center', fontsize=12, color='white')

# Adjust legend: 3 columns, place at bottom center of chart
ax.legend(ncol=3, loc='upper center', bbox_to_anchor=(0.47, -0.05), frameon=False)

# Rotate x-axis tick labels
# plt.xticks(rotation=40, ha='right')

# Save image (300 DPI high resolution)
plt.tight_layout()
plt.savefig(save_path, dpi=300)

# Display chart
plt.show()