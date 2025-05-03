"""
Chart engine that handles matplotlib initialization and figure creation
This separates matplotlib from the main application to prevent unwanted windows
"""
import matplotlib
matplotlib.use('Agg')  # Force non-interactive backend

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import re
from typing import Dict, List, Optional, Tuple, Union, Any
import logging
from datetime import datetime, timedelta

# Configure logger
logger = logging.getLogger(__name__)

DEBUG = True

def debug_print(message):
    if DEBUG:
        print(f"DEBUG [CHART]: {message}")

def process_excel_file(file_path: str) -> plt.Figure:
    """
    Process Excel file and create a Gantt chart.
    
    Args:
        file_path: Path to the Excel file
        
    Returns:
        Matplotlib figure object containing the Gantt chart
    """
    try:
        # Load the data
        df = pd.read_excel(file_path)
        
        # Identify relevant columns
        columns = detect_columns(df)
        
        # Generate the chart
        fig, _ = create_gantt_chart(df, columns)
        
        return fig
    except Exception as e:
        logger.error(f"Error processing Excel file: {str(e)}")
        raise

def detect_columns(df: pd.DataFrame) -> Dict[str, str]:
    """
    Detect relevant column names in the DataFrame.
    
    Args:
        df: DataFrame to analyze
        
    Returns:
        Dictionary mapping standard column names to actual column names
    """
    column_mapping = {}
    columns_lower = {col.lower(): col for col in df.columns}
    
    # Define patterns to search for
    patterns = {
        'task': ['task', 'name', 'description'],
        'duration': ['duration', 'weeks', 'days'],
        'phase': ['phase', 'category', 'group'],
        'start_date': ['start', 'begin'],
        'end_date': ['end', 'finish']
    }
    
    # Find matching columns
    for key, patterns_list in patterns.items():
        for pattern in patterns_list:
            for col_lower, col_original in columns_lower.items():
                if pattern in col_lower:
                    column_mapping[key] = col_original
                    break
            if key in column_mapping:
                break
    
    # Validate required columns
    required_cols = ['task', 'start_date', 'end_date']
    missing_cols = [col for col in required_cols if col not in column_mapping]
    
    if missing_cols:
        raise ValueError(f"Required columns not found: {', '.join(missing_cols)}")
    
    return column_mapping

def create_gantt_chart(
    df: pd.DataFrame, 
    columns: Dict[str, str],
    custom_colors: Optional[Dict[str, str]] = None,
    background_color: str = "#1f2937",
    grid_color: str = "#ffffff"
) -> Tuple[plt.Figure, plt.Axes]:
    """
    Create a Gantt chart from the processed data.
    
    Args:
        df: DataFrame containing the task data
        columns: Dictionary mapping standard column names to actual column names
        custom_colors: Optional dictionary mapping phases to custom colors
        background_color: Color for chart background
        grid_color: Color for grid lines and axes
        
    Returns:
        Tuple containing the figure and axes objects
    """
    # Extract column names
    task_col = columns['task']
    start_col = columns['start_date']
    end_col = columns['end_date']
    phase_col = columns.get('phase')
    duration_col = columns.get('duration')
    
    # Create figure and axis with custom background
    fig, ax = plt.subplots(1, 1, figsize=(12, 8), facecolor=background_color)
    ax.set_facecolor(background_color)
    
    # Process dates if they're not already datetime
    if not pd.api.types.is_datetime64_any_dtype(df[start_col]):
        df[start_col] = pd.to_datetime(df[start_col])
    if not pd.api.types.is_datetime64_any_dtype(df[end_col]):
        df[end_col] = pd.to_datetime(df[end_col])
    
    # Sort tasks by start date (ascending)
    df = df.sort_values(by=start_col)
    
    # Set up color mapping for phases
    colors = plt.cm.Dark2.colors
    if phase_col and phase_col in df.columns:
        phases = sorted(df[phase_col].unique())
        
        # Use custom colors if provided, otherwise use default color scheme
        if custom_colors:
            color_map = {
                phase: custom_colors.get(phase, colors[i % len(colors)]) 
                for i, phase in enumerate(phases)
            }
        else:
            color_map = {
                phase: colors[i % len(colors)] 
                for i, phase in enumerate(phases)
            }
    else:
        color_map = {}
        phases = []
    
    # Plot each task as a horizontal bar
    # Track task positions to ensure correct ordering
    y_positions = []
    y_ticks = []
    
    # Process tasks from earliest to latest
    for i, (_, row) in enumerate(df.iterrows()):
        task = row[task_col]
        start_date = row[start_col]
        end_date = row[end_col]
        
        # Ensure the task is positioned from top (earliest) to bottom (latest)
        y_pos = len(df) - 1 - i  # Reverse the position so earliest is at top
        y_positions.append(y_pos)
        y_ticks.append(task)
        
        # Calculate duration in days
        duration_days = (end_date - start_date).days
        if duration_days <= 0:
            duration_days = 1  # Ensure minimum visible duration
        
        # Get color based on phase
        if phase_col and phase_col in df.columns:
            phase = row[phase_col]
            color = color_map.get(phase, 'blue')
        else:
            color = 'steelblue'
        
        # Plot the task bar
        ax.barh(y_pos, duration_days, left=start_date, height=0.5, 
                color=color, edgecolor='black', alpha=0.8)
        
        # Add duration label
        if duration_col and duration_col in df.columns:
            duration_text = f"{row[duration_col]}"
        else:
            duration_text = f"{duration_days}d"
            
        # Center the text in the bar
        text_x = start_date + pd.Timedelta(days=duration_days/2)
        ax.text(text_x, y_pos, duration_text,
                ha='center', va='center', color='white', fontweight='bold')
    
    # Set up the axes with task names
    ax.set_yticks(y_positions)
    ax.set_yticklabels(y_ticks)
    
    # Format the x-axis for dates with custom grid color
    ax.grid(True, axis='x', alpha=0.3, color=grid_color)
    
    # Add week and month gridlines with custom color
    start = df[start_col].min()
    end = df[end_col].max()
    date_range = pd.date_range(start=start, end=end, freq='W')
    
    for date in date_range:
        ax.axvline(date, color=grid_color, linestyle='--', alpha=0.3)
    
    month_range = pd.date_range(start=start, end=end, freq='MS')
    for date in month_range:
        ax.axvline(date, color=grid_color, linestyle='-', alpha=0.2)
    
    # Rotate date labels for better readability
    fig.autofmt_xdate()
    
    # Dark theme styling with custom colors
    for spine in ax.spines.values():
        spine.set_color(grid_color)
    ax.tick_params(colors=grid_color)
    
    # Add title and labels with custom grid color
    ax.set_title('Project Timeline', color=grid_color, fontsize=16)
    ax.set_xlabel('Date', color=grid_color)
    
    # Create a legend with improved visibility if phases are available
    if phase_col and phase_col in df.columns and len(phases) > 0:
        handles = [plt.Rectangle((0,0), 1, 1, color=color_map[phase]) for phase in phases]
        legend = ax.legend(
            handles, 
            phases, 
            loc='upper right', 
            facecolor=background_color, 
            edgecolor=grid_color
        )
        # Make legend text white for better visibility
        for text in legend.get_texts():
            text.set_color(grid_color)
    
    plt.tight_layout()
    return fig, ax

def get_available_colormaps() -> List[str]:
    """
    Get a list of available color maps from matplotlib.
    
    Returns:
        List of color map names suitable for Gantt charts
    """
    # Select a subset of colormaps that work well for Gantt charts
    recommended_maps = [
        'Dark2', 'Set1', 'Set2', 'Set3', 'tab10', 'tab20', 
        'Paired', 'Accent', 'Pastel1', 'Pastel2'
    ]
    
    return recommended_maps

def generate_color_scheme(phases: List[str], colormap_name: str = 'Dark2') -> Dict[str, str]:
    """
    Generate a color scheme for phases using the specified colormap.
    
    Args:
        phases: List of phase names
        colormap_name: Name of matplotlib colormap to use
        
    Returns:
        Dictionary mapping phases to colors in hex format
    """
    cmap = plt.cm.get_cmap(colormap_name, max(len(phases), 8))
    
    # Convert colors to hex format
    colors = {}
    for i, phase in enumerate(phases):
        rgba = cmap(i % cmap.N)
        hex_color = '#{:02x}{:02x}{:02x}'.format(
            int(rgba[0] * 255), 
            int(rgba[1] * 255), 
            int(rgba[2] * 255)
        )
        colors[phase] = hex_color
    
    return colors

def extract_phases_from_file(file_path: str) -> List[str]:
    """
    Extract unique phase values from an Excel file.
    
    Args:
        file_path: Path to the Excel file
        
    Returns:
        List of unique phase values
    """
    try:
        df = pd.read_excel(file_path)
        columns = detect_columns(df)
        
        if 'phase' in columns and columns['phase'] in df.columns:
            phases = sorted(df[columns['phase']].unique().tolist())
            return phases
        else:
            return []
    except Exception as e:
        logger.error(f"Error extracting phases: {str(e)}")
        return []

def create_gantt_from_file(
    file_path: str, 
    custom_colors: Optional[Dict[str, str]] = None,
    background_color: str = "#1f2937",
    grid_color: str = "#ffffff"
) -> plt.Figure:
    """
    Main function to create a Gantt chart from an Excel file.
    
    Args:
        file_path: Path to the Excel file
        custom_colors: Optional dictionary mapping phases to custom colors
        background_color: Color for chart background
        grid_color: Color for grid lines and axes
        
    Returns:
        Matplotlib figure object containing the Gantt chart
    """
    logger.debug(f"Creating chart from file: {file_path}")
    try:
        # Load data once and reuse
        df = pd.read_excel(file_path)
        columns = detect_columns(df)
        
        # Generate the figure with the chart and custom colors
        fig, _ = create_gantt_chart(
            df, 
            columns, 
            custom_colors, 
            background_color, 
            grid_color
        )
        
        # Important: Set Matplotlib to use a backend that provides navigation
        plt.rcParams['toolbar'] = 'toolbar2'
        
        return fig
    except Exception as e:
        logger.debug(f"Error creating chart: {str(e)}")
        raise