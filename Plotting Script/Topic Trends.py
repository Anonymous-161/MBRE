import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.colors as mcolors
import colorsys
from matplotlib.path import Path
from matplotlib.patches import PathPatch
import os
from datetime import datetime

# Set Chinese font support
# plt.rcParams["font.family"] = ["Times New Roman", "serif"]
plt.rcParams['axes.unicode_minus'] = False  # Fix negative sign display issue


def read_and_process_data(file_path):
    """Read Excel file and process data"""
    # Read Excel file
    df = pd.read_excel(file_path)

    # Data cleaning: Remove null values in Publication Year or Topic columns
    df = df.dropna(subset=['Publication Year', 'Topic'])

    # Convert Publication Year to integer type
    df['Publication Year'] = df['Publication Year'].astype(int)

    # Create year range mapping - 5-year intervals (left-closed, right-open)
    year_bins = [2010, 2015, 2020, 2025]  # Intervals: [2010,2015), [2015,2020), [2020,2025)
    year_labels = [f"{start}-{end - 1}" for start, end in zip(year_bins[:-1], year_bins[1:])]

    # Map years to corresponding intervals
    df['Year Range'] = pd.cut(df['Publication Year'], bins=year_bins, labels=year_labels, right=False)

    # Remove data outside the specified year range
    df = df.dropna(subset=['Year Range'])

    # Count number of papers per topic in each year range
    topic_counts = df.groupby(['Year Range', 'Topic']).size().reset_index(name='Count')

    return topic_counts, year_labels


def generate_colors(num_colors):
    """Generate similar, low-key and sophisticated color scheme"""
    # Use different shades of the same color system
    base_hue = 210 / 360  # Blue color system
    base_saturation = 0.7

    colors = []
    for i in range(num_colors):
        # Adjust lightness and saturation on the same hue
        lightness = 0.9 - (i % 5) * 0.15  # Vary between 0.9-0.15
        saturation = base_saturation - (i // 5) * 0.1  # Vary between 0.7-0.4

        # Convert HLS to RGB using colorsys module
        r, g, b = colorsys.hls_to_rgb(base_hue, lightness, max(0.3, saturation))
        colors.append(mcolors.rgb2hex((r, g, b)))

    return colors


def create_flow_data(topic_counts, year_labels):
    """Create flow data"""
    # Get all unique topics
    all_topics = sorted(topic_counts['Topic'].unique())
    num_topics = len(all_topics)

    # Create mapping from year range to index
    year_to_idx = {year: i for i, year in enumerate(year_labels)}

    # Create topic-color mapping
    topic_colors = dict(zip(all_topics, generate_colors(num_topics)))

    # Initialize flow data structure
    flow_data = []

    # Create flow for each topic
    for topic in all_topics:
        # Get paper counts of the topic in each year range
        topic_data = topic_counts[topic_counts['Topic'] == topic].copy()

        # Add 0-count records for missing year ranges of the topic
        for year in year_labels:
            if year not in topic_data['Year Range'].values:
                topic_data = pd.concat([topic_data, pd.DataFrame({
                    'Year Range': [year],
                    'Topic': [topic],
                    'Count': [0]
                })])

        # Sort by year
        topic_data = topic_data.sort_values('Year Range')

        # Store flow data for the topic
        flow_data.append({
            'topic': topic,
            'counts': [topic_data[topic_data['Year Range'] == year]['Count'].values[0] for year in year_labels],
            'color': topic_colors[topic]
        })

    return flow_data, topic_colors, year_labels


def plot_topic_trend(flow_data, topic_colors, year_labels, output_path):
    """Plot topic trend chart"""
    # Calculate maximum paper count for chart height determination
    max_count = max(max(data['counts']) for data in flow_data)
    total_topics = len(flow_data)

    # Create canvas
    fig, ax = plt.subplots(figsize=(18, 11))  # Increase height to accommodate more topics
    plt.subplots_adjust(left=0.2, right=0.95, top=0.92, bottom=0.1)  # Adjust left margin

    # Define x-axis positions for year ranges
    x_positions = np.linspace(0, 1, len(year_labels))

    # Calculate starting position for each topic area
    topic_positions = {}
    current_y = 0

    for i, data in enumerate(flow_data):
        counts = data['counts']
        start_value = counts[0]
        topic_positions[data['topic']] = current_y + start_value / 2
        current_y += start_value

    # Plot flow paths
    for i, data in enumerate(flow_data):
        topic = data['topic']
        counts = data['counts']
        color = data['color']

        # Plot flow between year ranges for each topic
        for j in range(len(year_labels) - 1):
            x_start = x_positions[j]
            x_end = x_positions[j + 1]
            y_start = sum(d['counts'][j] for d in flow_data[:i]) + counts[j] / 2
            y_end = sum(d['counts'][j + 1] for d in flow_data[:i]) + counts[j + 1] / 2

            # Create Bezier curve path
            verts = [
                (x_start, y_start - counts[j] / 2),  # Start point
                (x_start + (x_end - x_start) / 3, y_start - counts[j] / 2),  # First control point
                (x_end - (x_end - x_start) / 3, y_end - counts[j + 1] / 2),  # Second control point
                (x_end, y_end - counts[j + 1] / 2),  # End point
                (x_end, y_end + counts[j + 1] / 2),  # End point
                (x_end - (x_end - x_start) / 3, y_end + counts[j + 1] / 2),  # Second control point
                (x_start + (x_end - x_start) / 3, y_start + counts[j] / 2),  # First control point
                (x_start, y_start + counts[j] / 2),  # Start point
                (x_start, y_start - counts[j] / 2),  # Closing point
            ]

            codes = [
                Path.MOVETO,
                Path.CURVE4,
                Path.CURVE4,
                Path.CURVE4,
                Path.LINETO,
                Path.CURVE4,
                Path.CURVE4,
                Path.CURVE4,
                Path.CLOSEPOLY,
            ]

            path = Path(verts, codes)
            patch = PathPatch(path, facecolor=color, alpha=0.8, edgecolor='none')
            ax.add_patch(patch)

    # Add year labels
    ax.set_xticks(x_positions)
    ax.set_xticklabels(year_labels, fontsize=12)

    # Set axis titles
    ax.set_xlabel('Publication year', fontsize=12)  # Set x-axis title
    ax.set_ylabel('Number of papers', fontsize=12)  # Set y-axis title

    # Add topic labels to the left of y-axis
    for topic, y_pos in topic_positions.items():
        color = topic_colors[topic]
        ax.text(-0.05, y_pos, topic, fontsize=10, ha='right', va='center',
                color='black', fontweight='bold', bbox=dict(facecolor=color, alpha=0.3, pad=2))

    # Set y-axis range
    ax.set_ylim(0, sum(d['counts'][0] for d in flow_data) + max_count * 0.1)

    # Set background and borders
    ax.set_facecolor('#f8f9fa')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['bottom'].set_visible(True)
    ax.spines['left'].set_visible(True)

    # Add grid lines
    ax.grid(axis='y', linestyle='--', alpha=0.7)

    # Save the chart
    plt.savefig(output_path, dpi=300, bbox_inches='tight')
    print(f"Chart saved to: {output_path}")


def main():
    """Main function"""
    print("Starting to generate literature topic trend chart...")

    # Fixed Excel file name
    file_path = "Data Table/Topic Trends.xlsx"

    # Check if file exists
    if not os.path.exists(file_path):
        print(f"Error: File '{file_path}' does not exist!")
        return

    try:
        # Read and process data
        print("Reading and processing data...")
        topic_counts, year_labels = read_and_process_data(file_path)

        # Create flow data
        print("Preparing visualization data...")
        flow_data, topic_colors, year_labels = create_flow_data(topic_counts, year_labels)

        # Generate output file name
        output_path = f"literature_topic_trend_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"

        # Plot topic trend chart
        print("Generating visualization chart...")
        plot_topic_trend(flow_data, topic_colors, year_labels, output_path)

        print("=" * 50)
        print("Literature topic trend visualization completed!")

    except Exception as e:
        print(f"Error occurred: {e}")


if __name__ == "__main__":
    main()