"""
Color selection dialog for QuickGantt application with support for saved themes.
"""

import tkinter as tk
from tkinter import ttk, simpledialog, messagebox  # Added messagebox here
from tkinter import colorchooser
from typing import Dict, List, Optional, Tuple, Any
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.colors as mcolors
import json
from pathlib import Path
import os

class ColorSelector:
    """
    Modal dialog for selecting colors for Gantt chart phases and backgrounds.
    Supports saving and loading custom color themes.
    """
    
    def __init__(self, parent: tk.Tk, phases: List[str], initial_colors: Optional[Dict[str, str]] = None, 
                 initial_background: Optional[str] = None, initial_grid: Optional[str] = None):
        """
        Initialize the color selector dialog.
        
        Args:
            parent: Parent tkinter window
            phases: List of phase names to assign colors to
            initial_colors: Optional dictionary of initial phase-color mappings
            initial_background: Optional initial background color
            initial_grid: Optional initial grid lines color
        """
        self.parent = parent
        self.phases = phases
        self.color_map = initial_colors or {}
        self.background_color = initial_background or "#1f2937"  # Default dark navy
        self.grid_color = initial_grid or "#ffffff"  # Default white grid
        self.result = None
        self.saved_themes = self.load_saved_themes()
        
        # Create dialog window as a proper modal dialog
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Chart Color Customization")
        self.dialog.geometry("550x750")  # Increased from 700 to 750 for better spacing
        
        # Make dialog modal
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.focus_set()
        
        # Prevent closing with X button - must use OK or Cancel
        self.dialog.protocol("WM_DELETE_WINDOW", self.cancel)
        
        # Center dialog on parent
        self.center_on_parent()
        
        # Main frame with notebook for tabs
        main_frame = ttk.Frame(self.dialog, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Create tabs
        phases_frame = ttk.Frame(notebook, padding="10")
        background_frame = ttk.Frame(notebook, padding="10")
        saved_themes_frame = ttk.Frame(notebook, padding="10")
        
        notebook.add(phases_frame, text="Phase Colors")
        notebook.add(background_frame, text="Background")
        notebook.add(saved_themes_frame, text="Saved Themes")
        
        # ====== Phase Colors Tab ======
        # Color theme selection
        theme_frame = ttk.LabelFrame(phases_frame, text="Color Theme", padding="10")
        theme_frame.pack(fill=tk.X, expand=False, pady=(0, 10))
        
        # Theme selector
        ttk.Label(theme_frame, text="Preset Color Themes:").grid(row=0, column=0, sticky=tk.W)
        
        # Get color maps
        self.color_maps = ["Dark2", "Set1", "Set2", "Paired", "tab10", "tab20", "Pastel1"]
        self.theme_var = tk.StringVar(value="Dark2")
        
        theme_combo = ttk.Combobox(theme_frame, textvariable=self.theme_var, values=self.color_maps)
        theme_combo.grid(row=0, column=1, sticky=tk.W, padx=(10, 0))
        theme_combo.bind("<<ComboboxSelected>>", self.apply_theme)
        
        apply_btn = ttk.Button(theme_frame, text="Apply Theme", command=self.apply_theme)
        apply_btn.grid(row=0, column=2, padx=10)
        
        # Preview
        preview_frame = ttk.LabelFrame(phases_frame, text="Preview", padding="10")
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        self.fig, self.ax = plt.subplots(figsize=(5, 2), tight_layout=True)
        self.canvas = FigureCanvasTkAgg(self.fig, master=preview_frame)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        # Color selector for each phase
        color_frame = ttk.LabelFrame(phases_frame, text="Custom Colors", padding="10")
        color_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create scrollable frame if there are many phases
        if len(phases) > 8:
            # Create canvas with scrollbar for many phases
            canvas = tk.Canvas(color_frame)
            scrollbar = ttk.Scrollbar(color_frame, orient="vertical", command=canvas.yview)
            scrollable_frame = ttk.Frame(canvas)
            
            scrollable_frame.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
            )
            
            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)
            
            canvas.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")
            
            target_frame = scrollable_frame
        else:
            target_frame = color_frame
        
        # Create color entry for each phase
        self.color_entries = {}
        for i, phase in enumerate(self.phases):
            ttk.Label(target_frame, text=f"{phase}:").grid(row=i, column=0, sticky=tk.W, pady=2)
            
            # Create a color display frame
            color_display = tk.Frame(target_frame, width=40, height=20, 
                                    bg=self.color_map.get(phase, "#cccccc"))
            color_display.grid(row=i, column=1, padx=10, pady=2)
            # Ensure the frame maintains its size
            color_display.grid_propagate(False)
            color_display.bind("<Button-1>", lambda e, p=phase: self.choose_color(p))
            
            # Store reference
            self.color_entries[phase] = color_display
        
        # ====== Background Tab ======
        bg_settings_frame = ttk.Frame(background_frame)
        bg_settings_frame.pack(fill=tk.X, expand=False, pady=(0, 10))
        
        # Background color
        ttk.Label(bg_settings_frame, text="Chart Background:").grid(row=0, column=0, sticky=tk.W, pady=10)
        
        self.bg_preview = tk.Frame(bg_settings_frame, width=40, height=20, bg=self.background_color)
        self.bg_preview.grid(row=0, column=1, padx=10, pady=10)
        self.bg_preview.grid_propagate(False)
        self.bg_preview.bind("<Button-1>", lambda e: self.choose_background_color())
        
        ttk.Button(bg_settings_frame, text="Change", 
                  command=self.choose_background_color).grid(row=0, column=2)
        
        # Grid lines color
        ttk.Label(bg_settings_frame, text="Grid Lines:").grid(row=1, column=0, sticky=tk.W, pady=10)
        
        self.grid_preview = tk.Frame(bg_settings_frame, width=40, height=20, bg=self.grid_color)
        self.grid_preview.grid(row=1, column=1, padx=10, pady=10)
        self.grid_preview.grid_propagate(False)
        self.grid_preview.bind("<Button-1>", lambda e: self.choose_grid_color())
        
        ttk.Button(bg_settings_frame, text="Change", 
                  command=self.choose_grid_color).grid(row=1, column=2)
        
        # Background color presets
        preset_frame = ttk.LabelFrame(background_frame, text="Background Presets", padding="10")
        preset_frame.pack(fill=tk.X, expand=False, pady=(10, 10))
        
        # Dark theme
        dark_btn = ttk.Button(preset_frame, text="Dark Theme", 
                             command=lambda: self.apply_preset("#1f2937", "#ffffff"))
        dark_btn.grid(row=0, column=0, padx=5, pady=5)
        
        # Light theme
        light_btn = ttk.Button(preset_frame, text="Light Theme", 
                              command=lambda: self.apply_preset("#f8f9fa", "#333333"))
        light_btn.grid(row=0, column=1, padx=5, pady=5)
        
        # Preview of background
        bg_preview_frame = ttk.LabelFrame(background_frame, text="Background Preview", padding="10")
        bg_preview_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        self.bg_fig, self.bg_ax = plt.subplots(figsize=(5, 3), tight_layout=True)
        self.bg_canvas = FigureCanvasTkAgg(self.bg_fig, master=bg_preview_frame)
        self.bg_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        # Save current theme button - moved to its own frame at the bottom for better visibility
        save_theme_frame = ttk.Frame(background_frame)
        save_theme_frame.pack(fill=tk.X, side=tk.BOTTOM, pady=(0, 5))
        
        save_theme_btn = ttk.Button(save_theme_frame, text="Save Current Theme", 
                                   command=self.save_current_theme)
        save_theme_btn.pack(pady=5)
        
        # ====== Saved Themes Tab ======
        saved_themes_label = ttk.Label(saved_themes_frame, 
                                     text="Select a saved theme to apply")
        saved_themes_label.pack(pady=(0, 10), anchor=tk.W)
        
        # Theme selection frame
        themes_list_frame = ttk.Frame(saved_themes_frame)
        themes_list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Create listbox with scrollbar
        scrollbar = ttk.Scrollbar(themes_list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.themes_listbox = tk.Listbox(themes_list_frame, 
                                        yscrollcommand=scrollbar.set,
                                        font=("Arial", 10),
                                        height=10)
        self.themes_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar.config(command=self.themes_listbox.yview)
        
        # Populate themes listbox
        self.populate_themes_listbox()
        
        # Buttons for saved themes
        themes_btn_frame = ttk.Frame(saved_themes_frame)
        themes_btn_frame.pack(fill=tk.X, pady=10)
        
        apply_theme_btn = ttk.Button(themes_btn_frame, text="Apply Selected Theme", 
                                    command=self.apply_saved_theme)
        apply_theme_btn.pack(side=tk.LEFT, padx=5)
        
        delete_theme_btn = ttk.Button(themes_btn_frame, text="Delete Selected Theme", 
                                     command=self.delete_saved_theme)
        delete_theme_btn.pack(side=tk.LEFT, padx=5)
        
        # Theme preview frame
        theme_preview_frame = ttk.LabelFrame(saved_themes_frame, text="Theme Preview")
        theme_preview_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.theme_preview_fig, self.theme_preview_ax = plt.subplots(figsize=(5, 3), tight_layout=True)
        self.theme_preview_canvas = FigureCanvasTkAgg(self.theme_preview_fig, master=theme_preview_frame)
        self.theme_preview_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        # Listen for theme selection changes
        self.themes_listbox.bind('<<ListboxSelect>>', self.preview_selected_theme)
        
        # Buttons at bottom for all tabs
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, expand=False, pady=(10, 10))  # Added padding at bottom
        
        cancel_btn = ttk.Button(button_frame, text="Cancel", command=self.cancel)
        cancel_btn.pack(side=tk.RIGHT, padx=5)
        
        ok_btn = ttk.Button(button_frame, text="OK", command=self.ok)
        ok_btn.pack(side=tk.RIGHT, padx=5)
        
        # Optional DPI awareness adjustment
        self.dialog.update_idletasks()  # Force layout update
        min_height = button_frame.winfo_reqheight() + main_frame.winfo_reqheight() + 50
        self.dialog.minsize(550, min_height)  # Set minimum size based on content
        
        # Initialize phase colors
        if not self.color_map:
            self.apply_theme()
        else:
            self.update_preview()
            
        # Initialize background preview
        self.update_bg_preview()
    
    def load_saved_themes(self) -> Dict[str, Dict[str, Any]]:
        """
        Load saved themes from the themes JSON file.
        
        Returns:
            Dictionary of theme names mapped to their settings
        """
        try:
            themes_path = self.get_themes_path()
            
            if not os.path.exists(themes_path):
                return {}
                
            with open(themes_path, 'r') as f:
                themes = json.load(f)
            
            return themes
        except Exception as e:
            print(f"Error loading saved themes: {e}")
            return {}
    
    def get_themes_path(self) -> str:
        """
        Get the path to the themes JSON file.
        
        Returns:
            Path to the themes file as a string
        """
        settings_dir = Path.home() / ".quickgantt"
        
        # Create directory if it doesn't exist
        if not settings_dir.exists():
            settings_dir.mkdir(exist_ok=True)
            
        return str(settings_dir / "color_themes.json")
    
    def save_themes(self) -> None:
        """Save all themes to the themes JSON file."""
        try:
            themes_path = self.get_themes_path()
            
            with open(themes_path, 'w') as f:
                json.dump(self.saved_themes, f, indent=2)
        except Exception as e:
            print(f"Error saving themes: {e}")
    
    def populate_themes_listbox(self) -> None:
        """Populate the themes listbox with saved theme names."""
        self.themes_listbox.delete(0, tk.END)
        
        # Sort theme names alphabetically
        theme_names = sorted(self.saved_themes.keys())
        
        for theme_name in theme_names:
            self.themes_listbox.insert(tk.END, theme_name)
    
    def save_current_theme(self) -> None:
        """
        Save the current color settings as a named theme.
        Prompts for confirmation if theme name already exists.
        """
        # Temporarily release grab for dialog
        self.dialog.grab_release()
        
        theme_name = simpledialog.askstring("Save Theme", 
                                          "Enter a name for this theme:",
                                          parent=self.dialog)
        
        # Re-establish modal behavior
        self.dialog.grab_set()
        self.dialog.focus_set()
        
        if not theme_name:
            return
            
        # Check if theme name already exists
        overwrite = True
        if (theme_name in self.saved_themes):
            # Import here to avoid circular imports
            from tkinter import messagebox
            
            # Release grab for confirmation dialog
            self.dialog.grab_release()
            
            overwrite = messagebox.askyesno(
                "Theme Already Exists",
                f"A theme named '{theme_name}' already exists. Do you want to overwrite it?",
                parent=self.dialog
            )
            
            # Re-establish modal behavior
            self.dialog.grab_set()
            self.dialog.focus_set()
            
            if not overwrite:
                return  # User chose not to overwrite, exit without saving
        
        # Create theme data
        theme_data = {
            'phase_colors': self.color_map.copy(),
            'background_color': self.background_color,
            'grid_color': self.grid_color
        }
        
        # Save to themes dictionary
        self.saved_themes[theme_name] = theme_data
        
        # Save to file
        self.save_themes()
        
        # Update listbox
        self.populate_themes_listbox()
        
        # Select the new theme
        idx = sorted(self.saved_themes.keys()).index(theme_name)
        self.themes_listbox.selection_clear(0, tk.END)
        self.themes_listbox.selection_set(idx)
        self.themes_listbox.see(idx)
        
        # Preview the theme
        self.preview_selected_theme(None)
    
    def apply_saved_theme(self) -> None:
        """Apply the selected theme from the listbox."""
        selection = self.themes_listbox.curselection()
        if not selection:
            return
            
        theme_name = self.themes_listbox.get(selection[0])
        theme_data = self.saved_themes.get(theme_name)
        
        if not theme_data:
            return
            
        # Apply theme data
        self.color_map = theme_data.get('phase_colors', {}).copy()
        self.background_color = theme_data.get('background_color', "#1f2937")
        self.grid_color = theme_data.get('grid_color', "#ffffff")
        
        # Update UI elements
        for phase, color in self.color_map.items():
            if phase in self.color_entries:
                self.color_entries[phase].configure(bg=color)
        
        self.bg_preview.configure(bg=self.background_color)
        self.grid_preview.configure(bg=self.grid_color)
        
        # Update previews
        self.update_preview()
        self.update_bg_preview()
    
    def delete_saved_theme(self) -> None:
        """Delete the selected theme from the listbox."""
        selection = self.themes_listbox.curselection()
        if not selection:
            return
            
        theme_name = self.themes_listbox.get(selection[0])
        
        # Temporarily release grab for dialog
        self.dialog.grab_release()
        
        confirm = messagebox.askyesno(
            "Confirm Delete",
            f"Are you sure you want to delete the theme '{theme_name}'?",
            parent=self.dialog
        )
        
        # Re-establish modal behavior
        self.dialog.grab_set()
        self.dialog.focus_set()
        
        if not confirm:
            return
            
        # Remove theme
        if theme_name in self.saved_themes:
            del self.saved_themes[theme_name]
            
        # Save to file
        self.save_themes()
        
        # Update listbox
        self.populate_themes_listbox()
        
        # Clear theme preview
        self.clear_theme_preview()
    
    def preview_selected_theme(self, event) -> None:
        """
        Preview the selected theme in the theme preview area.
        
        Args:
            event: Event data (can be None when called programmatically)
        """
        selection = self.themes_listbox.curselection()
        if not selection:
            self.clear_theme_preview()
            return
            
        theme_name = self.themes_listbox.get(selection[0])
        theme_data = self.saved_themes.get(theme_name)
        
        if not theme_data:
            self.clear_theme_preview()
            return
            
        # Update theme preview
        self.update_theme_preview(theme_data)
    
    def update_theme_preview(self, theme_data: Dict[str, Any]) -> None:
        """
        Update the theme preview with the specified theme data.
        
        Args:
            theme_data: Theme data containing colors and settings
        """
        self.theme_preview_ax.clear()
        
        # Get colors from theme
        bg_color = theme_data.get('background_color', "#1f2937")
        grid_color = theme_data.get('grid_color', "#ffffff")
        phase_colors = theme_data.get('phase_colors', {})
        
        # Set background color
        self.theme_preview_fig.patch.set_facecolor(bg_color)
        self.theme_preview_ax.set_facecolor(bg_color)
        
        # Calculate text color based on background
        r, g, b = mcolors.hex2color(bg_color)
        brightness = (r * 299 + g * 587 + b * 114) / 1000
        text_color = 'white' if brightness < 0.6 else 'black'
        
        # Create a simple preview chart
        labels = ["Task A", "Task B", "Task C", "Task D"]
        phases = ["Phase 1", "Phase 1", "Phase 2", "Phase 2"]
        
        # Sample tasks with different durations
        durations = [10, 8, 12, 6]
        starts = [1, 3, 6, 8]
        
        # Plot tasks in reverse order for top-to-bottom layout
        for i, (label, start, duration, phase) in enumerate(
            zip(labels, starts, durations, phases)
        ):
            # Calculate position so earliest tasks are at the top
            y_pos = len(labels) - 1 - i
            
            # Get color for this phase
            if phase in phase_colors:
                color = phase_colors[phase]
            else:
                # Fallback to default color if phase not in theme
                color = plt.cm.Dark2(i % 8)
            
            self.theme_preview_ax.barh(y_pos, duration, left=start, height=0.5, 
                                      color=color, edgecolor='black')
        
        # Add grid lines in theme grid color
        self.theme_preview_ax.grid(True, axis='x', alpha=0.3, color=grid_color)
        
        # Style axis elements with text color
        self.theme_preview_ax.set_yticks(range(len(labels)))
        self.theme_preview_ax.set_yticklabels(reversed(labels), color=text_color)
        self.theme_preview_ax.tick_params(colors=text_color)
        
        for spine in self.theme_preview_ax.spines.values():
            spine.set_color(grid_color)
        
        self.theme_preview_ax.set_title("Theme Preview", color=text_color)
        
        self.theme_preview_canvas.draw()
    
    def clear_theme_preview(self) -> None:
        """Clear the theme preview area."""
        self.theme_preview_ax.clear()
        self.theme_preview_ax.set_title("No theme selected")
        self.theme_preview_canvas.draw()
    
    def center_on_parent(self) -> None:
        """Center the dialog window on its parent."""
        self.dialog.update_idletasks()
        
        # Get parent and dialog dimensions
        parent_width = self.parent.winfo_width()
        parent_height = self.parent.winfo_height()
        parent_x = self.parent.winfo_rootx()
        parent_y = self.parent.winfo_rooty()
        
        dialog_width = self.dialog.winfo_width()
        dialog_height = self.dialog.winfo_height()
        
        # Calculate centered position
        x = parent_x + (parent_width - dialog_width) // 2
        y = parent_y + (parent_height - dialog_height) // 2
        
        # Set position
        self.dialog.geometry(f"+{max(0, x)}+{max(0, y)}")
    
    def apply_theme(self, event=None) -> None:
        """
        Apply the selected color theme to all phases.
        
        Args:
            event: Optional event data when triggered by combobox selection
        """
        theme = self.theme_var.get()
        cmap = plt.cm.get_cmap(theme, max(len(self.phases), 8))
        
        # Apply colors from the theme
        for i, phase in enumerate(self.phases):
            rgba = cmap(i % cmap.N)
            hex_color = '#{:02x}{:02x}{:02x}'.format(
                int(rgba[0] * 255), 
                int(rgba[1] * 255), 
                int(rgba[2] * 255)
            )
            self.color_map[phase] = hex_color
            self.color_entries[phase].configure(bg=hex_color)
        
        self.update_preview()
    
    def choose_color(self, phase: str) -> None:
        """
        Open color chooser for a specific phase.
        
        Args:
            phase: Phase name to choose color for
        """
        current_color = self.color_map.get(phase, "#cccccc")
        
        # Temporarily release grab so colorchooser dialog works properly
        self.dialog.grab_release()
        
        # Use the imported colorchooser directly
        color = colorchooser.askcolor(
            color=current_color,
            title=f"Select Color for {phase}",
            parent=self.dialog  # Set parent to ensure proper modal behavior
        )
        
        # Re-establish modal behavior
        self.dialog.grab_set()
        self.dialog.focus_set()
        
        if color[1]:  # If color is selected (not canceled)
            self.color_map[phase] = color[1]
            self.color_entries[phase].configure(bg=color[1])
            self.update_preview()
    
    def choose_background_color(self) -> None:
        """Open color chooser for chart background."""
        # Temporarily release grab
        self.dialog.grab_release()
        
        color = colorchooser.askcolor(
            color=self.background_color,
            title="Select Chart Background Color",
            parent=self.dialog
        )
        
        # Re-establish modal behavior
        self.dialog.grab_set()
        self.dialog.focus_set()
        
        if color[1]:
            self.background_color = color[1]
            self.bg_preview.configure(bg=color[1])
            self.update_bg_preview()
    
    def choose_grid_color(self) -> None:
        """Open color chooser for grid lines."""
        # Temporarily release grab
        self.dialog.grab_release()
        
        color = colorchooser.askcolor(
            color=self.grid_color,
            title="Select Grid Lines Color",
            parent=self.dialog
        )
        
        # Re-establish modal behavior
        self.dialog.grab_set()
        self.dialog.focus_set()
        
        if color[1]:
            self.grid_color = color[1]
            self.grid_preview.configure(bg=color[1])
            self.update_bg_preview()
    
    def apply_preset(self, bg_color: str, grid_color: str) -> None:
        """
        Apply a preset background and grid color combination.
        
        Args:
            bg_color: Background color in hex format
            grid_color: Grid lines color in hex format
        """
        self.background_color = bg_color
        self.grid_color = grid_color
        self.bg_preview.configure(bg=bg_color)
        self.grid_preview.configure(bg=grid_color)
        self.update_bg_preview()
    
    def update_preview(self) -> None:
        """Update the phase colors preview chart."""
        self.ax.clear()
        
        # Create a simple bar chart as preview
        y_pos = range(len(self.phases))
        
        for i, phase in enumerate(self.phases):
            color = self.color_map.get(phase, "#cccccc")
            self.ax.barh(i, 1, color=color, height=0.7, edgecolor='black')
            
            # Determine text color based on background brightness for better visibility
            r, g, b = mcolors.hex2color(color)
            brightness = (r * 299 + g * 587 + b * 114) / 1000
            text_color = 'white' if brightness < 0.6 else 'black'
            
            self.ax.text(0.5, i, phase, ha='center', va='center', 
                        color=text_color, fontweight='bold')
        
        self.ax.set_yticks([])
        self.ax.set_xticks([])
        self.ax.set_xlim(0, 1)
        self.ax.set_title("Phase Colors Preview")
        
        self.canvas.draw()
    
    def update_bg_preview(self) -> None:
        """
        Update the background preview chart with tasks displayed from top-left to bottom-right.
        
        This method creates a sample Gantt chart in the background preview panel
        with proper task ordering.
        """
        self.bg_ax.clear()
        
        # Set background color
        self.bg_fig.patch.set_facecolor(self.background_color)
        self.bg_ax.set_facecolor(self.background_color)
        
        # Create a sample Gantt chart
        labels = ["Task 1", "Task 2", "Task 3", "Task 4", "Task 5"]
        
        # Calculate text color based on background
        r, g, b = mcolors.hex2color(self.background_color)
        brightness = (r * 299 + g * 587 + b * 114) / 1000
        text_color = 'white' if brightness < 0.6 else 'black'
        
        # Sample tasks with different durations and start dates
        # These are arranged so that earlier tasks appear at the top
        durations = [10, 8, 12, 6, 9]
        starts = [1, 3, 6, 8, 12]
        
        # Plot tasks in reverse order to get the earliest task at the top
        for i, (label, start, duration) in enumerate(zip(labels, starts, durations)):
            # Calculate y-position so earliest tasks are at the top
            y_pos = len(labels) - 1 - i
            
            self.bg_ax.barh(y_pos, duration, left=start, height=0.5, 
                           color=plt.cm.Dark2(i % 8), edgecolor='black')
        
        # Add grid lines in selected color
        self.bg_ax.grid(True, axis='x', alpha=0.3, color=self.grid_color)
        
        # Style axis elements with our text color
        self.bg_ax.set_yticks(range(len(labels)))
        self.bg_ax.set_yticklabels(reversed(labels), color=text_color)  # Use reversed labels
        self.bg_ax.tick_params(colors=text_color)
        
        for spine in self.bg_ax.spines.values():
            spine.set_color(self.grid_color)
        
        self.bg_ax.set_title("Background Preview", color=text_color)
        self.bg_ax.set_xlabel("Timeline", color=text_color)
        
        self.bg_canvas.draw()
    
    def ok(self) -> None:
        """Save the color selections and close the dialog."""
        self.result = {
            'phase_colors': self.color_map.copy(),
            'background_color': self.background_color,
            'grid_color': self.grid_color
        }
        self.dialog.destroy()
    
    def cancel(self) -> None:
        """Cancel color selection and close the dialog."""
        self.result = None
        self.dialog.destroy()


def select_colors(parent: tk.Tk, phases: List[str], 
                 initial_colors: Optional[Dict[str, str]] = None,
                 initial_background: Optional[str] = None,
                 initial_grid: Optional[str] = None) -> Optional[Dict[str, Any]]:
    """
    Open a modal color selection dialog and return the chosen colors.
    
    Args:
        parent: Parent tkinter window
        phases: List of phase names to assign colors to
        initial_colors: Optional dictionary of initial phase-color mappings
        initial_background: Optional initial background color
        initial_grid: Optional initial grid lines color
        
    Returns:
        Dictionary with 'phase_colors', 'background_color', and 'grid_color' keys,
        or None if canceled
    """
    # Import here to avoid circular imports
    from tkinter import messagebox
    
    selector = ColorSelector(
        parent, 
        phases, 
        initial_colors, 
        initial_background, 
        initial_grid
    )
    
    parent.wait_window(selector.dialog)
    
    return selector.result