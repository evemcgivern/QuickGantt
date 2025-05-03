# Set backend once at the top, before ANY other imports
import matplotlib
matplotlib.use('Agg')  # Use non-interactive backend

import pandas as pd
from matplotlib import pyplot as plt
import tkinter as tk
from tkinter import filedialog
import numpy as np
import re
import matplotlib.dates as mdates
import matplotlib.font_manager as fm

# Add debug printing
DEBUG = True

def debug_print(message):
    if DEBUG:
        print(f"DEBUG [MAIN]: {message}")

# At the start of the file
debug_print("Importing main.py")
debug_print(f"Using matplotlib backend: {matplotlib.get_backend()}")

def main(root_window=None):
    """Main function that generates a Gantt chart from Excel data"""
    debug_print(f"main() called with root_window={root_window}")
    
    # Set the default font family to Arial
    plt.rcParams['font.family'] = 'Arial'
    
    debug_print("Using filedialog to get file path")
    # Ask user to select an Excel file using the provided root window
    file_path = filedialog.askopenfilename(
        parent=root_window,  # Use the passed window as parent
        title="Select Excel File",
        filetypes=[("Excel files", "*.xlsx;*.xls")]
    )
    debug_print(f"File path selected: {file_path}")
    
    # Check if a file was selected
    if not file_path:
        debug_print("No file selected, returning None")
        print("No file selected. Exiting program.")
        return None
    
    try:
        debug_print("Calling process_excel_file")
        fig = process_excel_file(file_path)
        debug_print("Returning figure")
        return fig
    except Exception as e:
        debug_print(f"Error in main: {e}")
        print(f"An error occurred: {e}")
        raise

def process_excel_file(file_path):
    debug_print("Start processing Excel file")
    # Read the Excel file
    df = pd.read_excel(file_path)
    
    # Print available columns for debugging
    print("Available columns in Excel file:")
    for col in df.columns:
        print(f"  - '{col}'")
    
    # Define required columns with case-insensitive matching
    required_column_patterns = {
        'Task': r'^task$',
        'Duration_Weeks': r'duration.*weeks',
        'Phase': r'^phase$',
        'Start_Date': r'start.*date',
        'End_Date': r'end.*date'
    }
    
    # Create a mapping of actual column names to standard names
    column_mapping = {}
    missing_columns = []
    
    for std_name, pattern in required_column_patterns.items():
        matching_cols = []
        for col in df.columns:
            # Use a direct equality check if possible
            if col.replace(' ', '_') == std_name:
                matching_cols = [col]
                break
            # Otherwise try case-insensitive regex pattern matching
            elif re.search(pattern, col, re.IGNORECASE):
                matching_cols.append(col)
        
        if matching_cols:
            column_mapping[std_name] = matching_cols[0]
        else:
            missing_columns.append(std_name)
    
    # Print the mapping for debugging
    print("\nColumn mapping:")
    for std_name, actual_name in column_mapping.items():
        print(f"  - '{std_name}' maps to '{actual_name}'")
    
    if missing_columns:
        print(f"\nError: Missing required columns: {', '.join(missing_columns)}")
        print(f"Available columns: {', '.join(df.columns)}")
        return
    
    # Rename columns to standard names for consistency (with underscores instead of spaces)
    df = df.rename(columns={v: k for k, v in column_mapping.items()})
    
    # Convert date columns to datetime
    df['Start_Date'] = pd.to_datetime(df['Start_Date'])
    df['End_Date'] = pd.to_datetime(df['End_Date'])
    
    # Define custom colors based on the provided image
    # Using a bright green, teal blue, and purple
    custom_colors = ['#90c144', '#00b2d4', '#9b59b6']
    
    # Create a color map for phases
    phases = df['Phase'].unique()
    
    # If there are more phases than colors, cycle through the colors
    phase_colors = {}
    for i, phase in enumerate(phases):
        phase_colors[phase] = custom_colors[i % len(custom_colors)]
    
    # Sort by start date first, then phase (earliest dates at the top)
    df = df.sort_values(['Start_Date', 'Phase'])
    
    # Create the plot
    fig, ax = plt.subplots(figsize=(12, 8))
    
    # Create lists for y-ticks and labels
    y_ticks = []
    y_labels = []
    
    # Plot each task as a horizontal bar with phase-based coloring
    for i, task in enumerate(df.itertuples()):
        # Use len(df)-i-1 to invert the y-axis (earliest at top)
        y_pos = len(df) - i - 1
        y_ticks.append(y_pos)
        y_labels.append(task.Task)
        
        # Calculate bar width in days
        bar_width = (task.End_Date - task.Start_Date).days
        
        # Plot the bar
        bar = ax.barh(y_pos, 
                bar_width, 
                left=task.Start_Date, 
                color=phase_colors[task.Phase],
                height=0.5)
        
        # Calculate midpoint of the bar for text placement
        bar_midpoint = task.Start_Date + pd.Timedelta(days=bar_width/2)
        
        # Add the duration weeks as text above the bar
        ax.text(bar_midpoint, y_pos + 0.25, 
                f"{task.Duration_Weeks}", 
                ha='center', va='bottom',
                fontweight='bold', fontsize=9,
                family='Arial',
                bbox=dict(facecolor='white', alpha=0.7, edgecolor='none', pad=1))
    
    # Add a legend for phases
    handles = [plt.Rectangle((0,0), 1, 1, color=color) for color in phase_colors.values()]
    legend = ax.legend(handles, phase_colors.keys(), title="Phases", loc="upper right")
    plt.setp(legend.get_texts(), family='Arial')
    plt.setp(legend.get_title(), family='Arial', color='white')
    
    # Format the x-axis to show dates by month
    ax.xaxis_date()
    
    # Set major ticks to be at the start of each month
    ax.xaxis.set_major_locator(mdates.MonthLocator())
    
    # Format the tick labels as month abbreviation and year (e.g., "Jan 2025")
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%b %Y'))
    
    # Set minor ticks at the start of each week for more granularity
    ax.xaxis.set_minor_locator(mdates.WeekdayLocator(byweekday=0))  # Monday as the first day of week
    
    # Set dark navy background color to match the image
    fig.set_facecolor('#1a1a44')
    ax.set_facecolor('#1a1a44')
    
    # Set labels with Arial font
    ax.set_xlabel('Date', color='white', fontfamily='Arial', fontsize=11)
    ax.set_ylabel('Task', color='white', fontfamily='Arial', fontsize=11)
    ax.set_title('Gantt Chart by Phase', color='white', fontfamily='Arial', fontsize=14, fontweight='bold')
    
    # Set y-axis to show all tasks with Arial font
    plt.yticks(y_ticks, y_labels, color='white', fontfamily='Arial')
    
    # Rotate date labels for better readability with Arial font
    plt.xticks(rotation=45, color='white', fontfamily='Arial')
    
    # Add grid lines for better readability
    ax.grid(True, axis='x', linestyle='--', alpha=0.3, color='white')
    
    # Add minor grid lines for weeks
    ax.grid(True, axis='x', which='minor', linestyle=':', alpha=0.2, color='white')
    
    # Set spine colors to white
    for spine in ax.spines.values():
        spine.set_color('white')
    
    # Set tick colors to white
    ax.tick_params(axis='x', colors='white')
    ax.tick_params(axis='y', colors='white')
    
    # Make sure legend text is white
    for text in legend.get_texts():
        text.set_color('white')
    
    plt.tight_layout()
    # Return the figure instead of showing it
    debug_print("Returning figure from process_excel_file")
    return fig  # Return the figure object instead of calling plt.show()

if __name__ == "__main__":
    debug_print("Running main.py directly")
    fig = main()
    if fig:
        debug_print("Showing figure with plt.show()")
        plt.show()
