import tkinter as tk; tk._default_root = None  # Prevent automatic Tk() creation

import os
os.environ['MPLBACKEND'] = 'Agg'  # Force matplotlib to use Agg backend globally
import tkinter as tk
from tkinter import messagebox, filedialog, ttk
import sys
import subprocess
import tempfile
import platform
import pandas as pd
import json
from pathlib import Path

_NavigationToolbar2Tk = None

# Add debug printing
DEBUG = True

def debug_print(message):
    if DEBUG:
        print(f"DEBUG [APP]: {message}")

# At the start of the file
debug_print("Starting app.py")

# Only import matplotlib-related modules and main.py after Tk is initialized
# This prevents them from creating their own Tk instances
_chart_engine = None
_plt = None
_FigureCanvasTkAgg = None
_NavigationToolbar2Tk = None

def lazy_imports():
    """Load matplotlib and other modules only when needed"""
    global _chart_engine, _FigureCanvasTkAgg, _NavigationToolbar2Tk
    if _chart_engine is None:
        debug_print("Performing lazy imports")
        import matplotlib
        matplotlib.use('Agg')  # Use non-interactive backend
        from matplotlib import pyplot as plt
        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
        import chart_engine
        _chart_engine = chart_engine
        _plt = plt
        _FigureCanvasTkAgg = FigureCanvasTkAgg
        _NavigationToolbar2Tk = NavigationToolbar2Tk
        debug_print("Lazy imports completed")

def use_system_appearance():
    debug_print("Setting system appearance")
    """Ensure the application uses the system's native appearance controls"""
    if platform.system() == "Windows":
        try:
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(1)  # Make the application DPI-aware
            
            # Use system theme for controls
            import tkinter.ttk as ttk
            style = ttk.Style()
            style.theme_use('winnative')  # Use Windows native theme
        except Exception:
            pass  # Ignore any errors if this fails

# Call this function before creating the main window
use_system_appearance()

# Determine if the application is running as a standalone executable or as a script
if getattr(sys, 'frozen', False):
    # If running as compiled executable
    application_path = os.path.dirname(sys.executable)
else:
    # If running as script
    application_path = os.path.dirname(os.path.abspath(__file__))

def create_sample_file():
    """Create and open a temporary sample Excel file"""
    # Create a temporary file that will be deleted when closed
    temp_dir = tempfile.gettempdir()
    sample_file_path = os.path.join(temp_dir, "QuickGantt_Sample.xlsx")
    
    # Sample data based on the image
    data = {
        'Task': ["Task 1", "Task 2", "Task 3", "Task 4", "Task 5", 
                "Task 6", "Task 7", "Task 8", "Task 9", 
                "Task 1", "Task 2", "Task 3", "Task 4", "Task 5"],
                
        'Duration (weeks)': [2, 2, 3, 3, 5, 6, 4, 4, 3.5, 6, 8, 10, 9, 9],
        
        'Phase': ['Phase 1', 'Phase 1', 'Phase 1', 'Phase 1', 'Phase 1',
                 'Phase 1', 'Phase 1', 'Phase 1', 'Phase 1',
                 'Phase 2', 'Phase 2', 'Phase 2', 'Phase 2', 'Phase 2'],
                 
        'Start Date': [
            '5/19/2025', '5/19/2025', '6/2/2025', '6/2/2025', '6/23/2025',
            '7/28/2025', '8/25/2025', '6/2/2025', '9/29/2025', '9/29/2025',
            '11/10/2025', '1/5/2026', '3/23/2026', '6/8/2026'
        ],
        
        'End Date': [
            '6/2/2025', '6/2/2025', '6/23/2025', '6/23/2025', '7/28/2025',
            '9/8/2025', '9/22/2025', '6/30/2025', '10/22/2025', '11/10/2025',
            '1/5/2026', '3/16/2026', '5/25/2026', '8/10/2026'
        ]
    }
    
    # Create DataFrame and save to Excel
    df = pd.DataFrame(data)
    df.to_excel(sample_file_path, index=False)
    
    # Open the file with default application
    try:
        if os.name == 'nt':  # Windows
            os.startfile(sample_file_path)
            # Show a message to the user about saving
            messagebox.showinfo(
                "Sample File Created", 
                "A sample file has been opened in Excel.\n\n"
                "Please use 'Save As...' to save it to your preferred location before making changes."
            )
        elif os.name == 'posix':  # macOS, Linux
            subprocess.call(['open', sample_file_path])
            messagebox.showinfo(
                "Sample File Created", 
                "A sample file has been opened.\n\n"
                "Please use 'Save As...' to save it to your preferred location before making changes."
            )
        else:
            messagebox.showinfo(
                "Sample File Created", 
                f"Sample file created at: {sample_file_path}\n\n"
                "This is a temporary file. Please save it to your preferred location."
            )
    except Exception as e:
        messagebox.showerror("Error", f"Could not open the sample file: {e}")

class QuickGanttApp:
    def __init__(self, root=None):
        """
        Initialize the QuickGantt application.
        
        Args:
            root: Optional tkinter root window. If not provided, a new one will be created.
        """
        if root is None:
            self.root = tk.Tk()
        else:
            self.root = root
            
        self.root.title("QuickGantt")
        
        # More compact initial size
        self.root.geometry("450x350")
        
        # Set minimum size to prevent window from being too small
        self.root.minsize(400, 300)
        
        # Set application icon
        setup_app_icon(self.root)
        
        # Track window closing events
        self.root.protocol("WM_DELETE_WINDOW", self.on_main_window_close)
        
        # Ensure we're using standard Windows decorations
        if hasattr(self.root, 'attributes'):
            self.root.attributes('-alpha', 1.0)  # Ensure window is fully opaque
            
        # Windows specific - make sure we're using standard decorations
        if os.name == 'nt':
            self.root.overrideredirect(False)  # This disables custom window borders
            
        self.setup_ui()
        
        # Configure the app to respect system theme (light/dark mode)
        self.check_system_theme()
        
    def check_system_theme(self):
        """Check if system is using dark mode and adjust UI accordingly"""
        try:
            if os.name == 'nt':  # Windows
                import winreg
                registry = winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER)
                key = winreg.OpenKey(registry, r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize")
                value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
                
                if value == 0:  # Dark mode
                    self.use_dark_theme()
                else:  # Light mode
                    self.use_light_theme()
            else:
                # Default to light theme for other systems
                self.use_light_theme()
        except Exception as e:
            # If any error occurs, default to light theme
            debug_print(f"Error detecting system theme: {e}")
            self.use_light_theme()

    def use_dark_theme(self):
        """
        Apply dark theme colors to the application.
        
        Note: When using ttk widgets, we need to use the ttk Style system
        rather than directly configuring widget colors.
        """
        try:
            # Create and configure ttk style for dark theme
            style = ttk.Style()
            
            # Configure TFrame
            style.configure('TFrame', background='#2d2d2d')
            
            # Configure TLabel
            style.configure('TLabel', background='#2d2d2d', foreground='white')
            
            # Configure TButton
            style.configure('TButton', background='#404040', foreground='white')
            style.map('TButton',
                     background=[('active', '#505050')],
                     foreground=[('active', 'white')])
            
            # Set window background color
            self.root.configure(bg='#2d2d2d')
            
            debug_print("Dark theme applied")
        except Exception as e:
            debug_print(f"Error applying dark theme: {e}")

    def use_light_theme(self):
        """
        Apply light theme colors to the application.
        
        Note: When using ttk widgets, we need to use the ttk Style system
        rather than directly configuring widget colors.
        """
        try:
            # Create and configure ttk style for light theme
            style = ttk.Style()
            
            # Configure TFrame
            style.configure('TFrame', background='#f0f0f0')
            
            # Configure TLabel
            style.configure('TLabel', background='#f0f0f0', foreground='black')
            
            # Configure TButton
            style.configure('TButton', background='#e0e0e0', foreground='black')
            style.map('TButton',
                    background=[('active', '#d0d0d0')],
                    foreground=[('active', 'black')])
            
            # Set window background color
            self.root.configure(bg='#f0f0f0')
            
            debug_print("Light theme applied")
        except Exception as e:
            debug_print(f"Error applying light theme: {e}")
        
    def setup_ui(self):
        """Set up the user interface elements."""
        # Use a main frame with proper padding
        main_frame = ttk.Frame(self.root, padding="20 20 20 20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # App title with proper styling
        title_label = ttk.Label(
            main_frame, 
            text="QuickGantt", 
            font=("Arial", 18, "bold")
        )
        title_label.pack(pady=(0, 15))
        
        # Description 
        desc_label = ttk.Label(
            main_frame,
            text="Generate Gantt charts from Excel data",
            font=("Arial", 10)
        )
        desc_label.pack(pady=(0, 20))
        
        # Button frame for centered layout
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)
        
        # Configure button frame's columns for centering
        button_frame.columnconfigure(0, weight=1)  # Left padding
        button_frame.columnconfigure(3, weight=1)  # Right padding
        
        # Create Gantt Chart button
        generate_btn = ttk.Button(
            button_frame, 
            text="Create Gantt Chart", 
            command=self.generate_chart,
            width=20
        )
        generate_btn.grid(row=0, column=1, pady=10, padx=5)
        
        # Create Sample Data button
        sample_btn = ttk.Button(
            button_frame,
            text="Create Sample Data",
            command=create_sample_file,
            width=20
        )
        sample_btn.grid(row=1, column=1, pady=10, padx=5)
        
        # Exit button
        exit_btn = ttk.Button(
            button_frame,
            text="Exit",
            command=self.on_main_window_close,
            width=20
        )
        exit_btn.grid(row=2, column=1, pady=10, padx=5)
        
        # Version info at bottom
        version_label = ttk.Label(
            main_frame,
            text="QuickGantt v1.0",
            font=("Arial", 8)
        )
        version_label.pack(side=tk.BOTTOM, pady=(20, 0))
        
        # Make the UI responsive to window resizing
        self.root.update_idletasks()
        
    def generate_chart(self):
        """
        Launch the chart generation process with optional color customization.
        
        Creates a new window to display the Gantt chart using matplotlib
        with a navigation toolbar.
        """
        try:
            debug_print("Starting generate_chart")
            lazy_imports()
            
            # Get the Excel file path using tkinter's file dialog
            file_path = filedialog.askopenfilename(
                parent=self.root,
                title="Select Excel File",
                filetypes=[("Excel files", "*.xlsx;*.xls")]
            )
            
            if not file_path:
                debug_print("No file selected")
                return
            
            # Extract phases from the file for color selection
            phases = _chart_engine.extract_phases_from_file(file_path)

            # Always show the color customization dialog now, instead of asking first
            from color_selector import select_colors

            # Get current chart settings (if any)
            current_settings = self.get_saved_color_settings() # This function should be defined to retrieve saved settings

            # Open color selector dialog
            color_settings = select_colors(
                self.root, 
                phases,
                initial_colors=current_settings.get('phase_colors') if current_settings else None,
                initial_background=current_settings.get('background_color') if current_settings else None,
                initial_grid=current_settings.get('grid_color') if current_settings else None
            )

            # If user canceled (clicked Cancel button OR closed the window), don't proceed
            if color_settings is None:
                debug_print("Color selection canceled, not generating chart")
                return  # Exit the function without generating a chart

            # If we get here, user clicked OK, so save the settings and generate chart
            self.save_color_settings(color_settings)

            # Now proceed with chart generation using the selected settings
            debug_print("Generating chart with custom colors")
            fig = _chart_engine.create_gantt_from_file(
                file_path, 
                color_settings['phase_colors'],
                color_settings['background_color'],
                color_settings['grid_color']
            )
            
            if fig:
                debug_print("Creating chart window")
                # Create a new top-level window for the chart
                chart_window = tk.Toplevel(self.root)
                chart_window.title("Gantt Chart")
                chart_window.geometry("1200x800")  # Set a default size
                
                # Configure chart window to expand properly
                chart_window.columnconfigure(0, weight=1)
                chart_window.rowconfigure(0, weight=1)
                chart_window.rowconfigure(1, weight=0)
                
                # Track chart window closing
                chart_window.protocol("WM_DELETE_WINDOW", 
                                     lambda: self.on_chart_window_close(chart_window))
                
                # Create a frame to hold everything - using grid
                chart_frame = tk.Frame(chart_window)
                chart_frame.grid(row=0, column=0, sticky="nsew")
                chart_frame.columnconfigure(0, weight=1)
                chart_frame.rowconfigure(0, weight=1)
                
                debug_print("Creating FigureCanvasTkAgg")
                # Create a canvas for the figure
                canvas = _FigureCanvasTkAgg(fig, master=chart_frame)
                canvas_widget = canvas.get_tk_widget()
                canvas_widget.grid(row=0, column=0, sticky="nsew")
                
                # Add toolbar - also using grid within the toolbar frame
                debug_print("Adding toolbar")
                toolbar_frame = tk.Frame(chart_window)
                toolbar_frame.grid(row=1, column=0, sticky="ew")
                
                # Use the globally imported NavigationToolbar2Tk
                # Creating a custom toolbar class that uses grid instead of pack
                class GridNavigationToolbar(_NavigationToolbar2Tk):
                    """
                    Custom navigation toolbar that uses grid layout and customized tooltips.
                    
                    This toolbar inherits from NavigationToolbar2Tk but modifies the tooltips
                    to be more user-friendly for Gantt charts and adds custom functionality
                    for printing and customizing the chart appearance.
                    """
                    # Override with only the standard buttons we know exist in matplotlib
                    toolitems = [
                        ('Home', 'Reset chart to original view', 'home', 'home'),
                        ('Back', 'Back to previous view', 'back', 'back'),
                        ('Forward', 'Forward to next view', 'forward', 'forward'),
                        ('Pan', 'Pan with left mouse, zoom with right', 'move', 'pan'),
                        ('Zoom', 'Zoom to rectangle', 'zoom_to_rect', 'zoom'),
                        ('Subplots', 'Configure chart margins', 'subplots', 'configure_subplots'),
                        ('Save', 'Save the chart', 'filesave', 'save_figure')
                    ]
                    
                    def __init__(self, canvas, parent, chart_window=None):
                        """
                        Initialize the custom navigation toolbar with additional custom buttons.
                        
                        Args:
                            canvas: The matplotlib canvas widget
                            parent: The parent tkinter widget
                            chart_window: The toplevel window containing the chart
                        """
                        # Initialize with standard buttons first
                        super().__init__(canvas, parent)
                        
                        # Store reference to parent window directly rather than trying to access through manager
                        self.parent_window = chart_window  # Use the explicitly passed chart_window
                        
                        # Add spacer (separating standard and custom buttons)
                        separator = tk.Frame(self, width=8, height=24, bd=0)
                        separator.pack(side=tk.LEFT, padx=5)
                        
                        # Add custom text buttons that don't need image files
                        self.print_button = tk.Button(
                            master=self, 
                            text="Print",
                            command=self.print_figure
                        )
                        self.print_button.pack(side=tk.LEFT, padx=2)
                        
                        self.customize_button = tk.Button(
                            master=self, 
                            text="Customize",
                            command=self.customize_chart
                        )
                        self.customize_button.pack(side=tk.LEFT, padx=2)
                        
                        # Add tooltips to custom buttons if possible
                        if hasattr(self, '_make_balloon'):
                            self._make_balloon()
                            if hasattr(self, '_balloon'):
                                self._balloon.bind(self.print_button, "Print the chart")
                                self._balloon.bind(self.customize_button, "Customize chart appearance")
                    
                    def customize_chart(self):
                        """
                        Open the color customization dialog to modify the chart appearance.
                        
                        This allows users to change colors without generating a new chart.
                        """
                        try:
                            # Get the app instance from the parent window
                            app = None
                            for widget in self.parent_window.winfo_children():
                                if hasattr(widget, 'winfo_children'):
                                    for child in widget.winfo_children():
                                        if isinstance(child, tk.Frame) and hasattr(child, 'winfo_toplevel'):
                                            app = child.winfo_toplevel().master
                                            break
                            
                            # If we can't find the app instance, try another approach
                            if not app or not hasattr(app, 'get_saved_color_settings'):
                                app = self.parent_window.master
                            
                            # Import required modules
                            from color_selector import select_colors
                            import matplotlib.colors as mcolors
                            
                            # Get current chart figure and extract data
                            fig = self.canvas.figure
                            
                            # Extract phases from the current chart
                            phases = []
                            background_color = fig.get_facecolor()
                            # Convert background color to hex format
                            background_color = convert_color_to_hex(background_color)
                            grid_color = "#ffffff"  # Default
                            
                            # Extract phases and colors from the current figure
                            for ax in fig.get_axes():
                                # Find grid color from existing chart
                                if ax.xaxis.get_gridlines() and len(ax.xaxis.get_gridlines()) > 0:
                                    grid_color = ax.xaxis.get_gridlines()[0].get_color()
                                    grid_color = convert_color_to_hex(grid_color)
                                
                                # Extract phase information from patches (bars)
                                for patch in ax.patches:
                                    # Try to get the label (phase) from patch properties
                                    # In Gantt charts, usually stored in the patch's custom properties
                                    if hasattr(patch, 'get_label') and patch.get_label():
                                        phase = patch.get_label()
                                        if phase and phase not in phases and phase != "_nolegend_":
                                            phases.append(phase)
                            
                            # If we couldn't extract phases from the chart, use defaults
                            if not phases:
                                phases = ["Phase 1", "Phase 2"]
                            
                            # Get current settings
                            current_settings = None
                            if hasattr(app, 'get_saved_color_settings'):
                                current_settings = app.get_saved_color_settings()
                            
                            # Open color selector dialog
                            color_settings = select_colors(
                                self.parent_window, 
                                phases,
                                initial_colors=current_settings.get('phase_colors') if current_settings else None,
                                initial_background=background_color,
                                initial_grid=grid_color
                            )
                            
                            # If user canceled, don't make changes
                            if color_settings is None:
                                return
                            
                            # Save the settings
                            if hasattr(app, 'save_color_settings'):
                                app.save_color_settings(color_settings)
                            
                            # Apply the new colors to the existing chart
                            self._apply_colors_to_chart(color_settings)
                            
                        except Exception as e:
                            # Log error but don't crash
                            print(f"DEBUG [APP]: Error in customize_chart: {e}")
                            import traceback
                            traceback.print_exc()
                            messagebox.showerror("Error", f"An error customizing chart: {e}")
                    
                    def _apply_colors_to_chart(self, color_settings: dict) -> None:
                        """
                        Apply the selected colors to the existing chart.
                        
                        Args:
                            color_settings: Dictionary containing color settings to apply
                        """
                        try:
                            import matplotlib.colors as mcolors
                            fig = self.canvas.figure
                            
                            # Apply background color - ensure it's in the right format for matplotlib
                            bg_color = color_settings['background_color']
                            fig.set_facecolor(bg_color)
                            
                            # Apply colors to each element
                            for ax in fig.get_axes():
                                # Set axis background
                                ax.set_facecolor(bg_color)
                                
                                # Set grid color
                                grid_color = color_settings['grid_color']
                                ax.grid(color=grid_color, alpha=0.3)
                                
                                # Calculate text color based on background
                                r, g, b = mcolors.to_rgb(bg_color)
                                brightness = (r * 299 + g * 587 + b * 114) / 1000
                                text_color = 'white' if brightness < 0.6 else 'black'
                                
                                # Update text colors
                                ax.title.set_color(grid_color)
                                ax.xaxis.label.set_color(grid_color)
                                ax.yaxis.label.set_color(grid_color)
                                
                                for label in ax.get_xticklabels() + ax.get_yticklabels():
                                    label.set_color(grid_color)
                                
                                # Update spines
                                for spine in ax.spines.values():
                                    spine.set_color(grid_color)
                                
                                # Update tick colors
                                ax.tick_params(colors=grid_color)
                                
                                # Update phase colors
                                phase_colors = color_settings.get('phase_colors', {})
                                for patch in ax.patches:
                                    if hasattr(patch, 'get_label') and patch.get_label() in phase_colors:
                                        patch.set_color(phase_colors[patch.get_label()])
                            
                            # Redraw the canvas to show changes
                            self.canvas.draw()
                            
                        except Exception as e:
                            print(f"DEBUG [APP]: Error applying colors: {e}")
                            import traceback
                            traceback.print_exc()
                    
                    def print_figure(self):
                        """
                        Print the chart using the platform's native print dialog.
                        
                        On Windows, this uses the built-in print functionality.
                        On other platforms, it saves to a temporary PDF and opens it.
                        """
                        try:
                            if platform.system() == "Windows":
                                # Use native Windows printing
                                import tempfile
                                import os
                                
                                # Create a temporary directory
                                temp_dir = tempfile.gettempdir()
                                temp_file = os.path.join(temp_dir, "quickgantt_print.png")
                                
                                # Save figure as high-resolution PNG
                                self.canvas.figure.savefig(
                                    temp_file, 
                                    format="png", 
                                    dpi=300,  # High resolution for printing
                                    bbox_inches="tight"
                                )
                                
                                # Open the native print dialog
                                try:
                                    # Try standard print dialog
                                    os.startfile(temp_file, "print")
                                except Exception:
                                    # If startfile fails, show a message with file location
                                    messagebox.showinfo(
                                        "Print Chart", 
                                        f"The chart has been saved as a PNG at:\n{temp_file}\n\n"
                                        "You can open this file and print it from any image viewer."
                                    )
                            
                            elif platform.system() == "Darwin":  # macOS
                                # On macOS, save to PDF and open with Preview
                                import tempfile
                                import os
                                import subprocess
                                
                                # Create temp PDF file
                                temp_file = tempfile.NamedTemporaryFile(
                                    suffix='.pdf',
                                    delete=False
                                ).name
                                
                                # Save figure as PDF
                                self.canvas.figure.savefig(
                                    temp_file, 
                                    format="pdf", 
                                    dpi=300,
                                    bbox_inches="tight"
                                )
                                
                                # Open with default PDF viewer which will have print option
                                subprocess.call(['open', temp_file])
                            
                            else:  # Linux or other
                                # For Linux, use xdg-open with a PDF
                                import tempfile
                                import subprocess
                                
                                # Create temp PDF file
                                temp_file = tempfile.NamedTemporaryFile(
                                    suffix='.pdf',
                                    delete=False
                                ).name
                                
                                # Save figure as PDF
                                self.canvas.figure.savefig(
                                    temp_file, 
                                    format="pdf", 
                                    dpi=300,
                                    bbox_inches="tight"
                                )
                                
                                # Try to open PDF with default viewer
                                try:
                                    subprocess.call(['xdg-open', temp_file])
                                except Exception:
                                    # Show message with file location if automatic opening fails
                                    messagebox.showinfo(
                                        "Print Chart", 
                                        f"The chart has been saved as a PDF at:\n{temp_file}\n\n"
                                        "You can open this file and print it from your PDF viewer."
                                    )
                                
                        except Exception as e:
                            print(f"DEBUG [APP]: Error in print_figure: {e}")
                            # Fall back to matplotlib's built-in print
                            messagebox.showinfo(
                                "Print Chart", 
                                "An error occurred with the print function. The chart has been saved as an image instead."
                            )
                            self.save_figure()  # Use the save dialog as a fallback
                
                # Use our modified toolbar class, passing the chart_window explicitly
                toolbar = GridNavigationToolbar(canvas, toolbar_frame, chart_window)
                toolbar.update()
                
                def ensure_window_on_top(self, window: tk.Toplevel) -> None:
                    """
                    Ensure that the specified window appears on top and has focus.
                    
                    This uses multiple methods to try to bring the window to the front
                    across different platforms and window managers.
                    
                    Args:
                        window: The Toplevel window to bring to the front
                    """
                    # First, ensure all pending events are processed
                    window.update_idletasks()
                    
                    # Use multiple methods to bring window to front
                    window.attributes('-topmost', True)  # Make window topmost temporarily
                    window.lift()  # Raise the window in the stacking order
                    window.focus_force()  # Force focus to this window
                    window.grab_set()  # Make window modal
                    
                    # After a brief delay, disable topmost to allow other windows to go in front if needed
                    def disable_topmost():
                        window.attributes('-topmost', False)
                        window.grab_release()  # Release the modal state
                    
                    # Schedule topmost attribute to be turned off after window appears
                    window.after(500, disable_topmost)
                    
                    # Additional platform-specific approaches
                    if platform.system() == "Windows":
                        try:
                            # Windows-specific: try using win32gui if available
                            import ctypes
                            ctypes.windll.user32.SetForegroundWindow(window.winfo_id())
                        except Exception:
                            debug_print("Windows-specific window activation failed, using standard methods")
                    
                    # One more update to ensure changes take effect
                    window.update()
                
        except Exception as e:
            debug_print(f"Error in generate_chart: {e}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"An error occurred: {e}")

    def on_main_window_close(self):
        """Handle main window closing event"""
        debug_print("Main window closing")
        self.root.destroy()

    def on_chart_window_close(self, window):
        """Handle chart window closing event"""
        debug_print("Chart window closing")
        window.destroy()

    def get_saved_color_settings(self) -> dict:
        """
        Retrieve saved color settings from a JSON file.
        
        Returns:
            Dictionary containing color settings or None if no saved settings exist
        """
        try:
            # Define path for settings file in user's home directory
            settings_dir = Path.home() / ".quickgantt"
            settings_file = settings_dir / "color_settings.json"
            
            # Create directory if it doesn't exist
            if not settings_dir.exists():
                settings_dir.mkdir(exist_ok=True)
                
            # Check if settings file exists
            if not settings_file.exists():
                return None
                
            # Read and parse settings file
            with open(settings_file, 'r') as f:
                settings = json.load(f)
                
            return settings
        except Exception as e:
            # Log error but don't crash the application
            print(f"DEBUG [APP]: Error loading saved color settings: {e}")
            return None

    def save_color_settings(self, settings: dict) -> None:
        """
        Save color settings to a JSON file for future use.
        
        Args:
            settings: Dictionary containing color settings to save
        """
        try:
            # Define path for settings file in user's home directory
            settings_dir = Path.home() / ".quickgantt"
            settings_file = settings_dir / "color_settings.json"
            
            # Create directory if it doesn't exist
            if not settings_dir.exists():
                settings_dir.mkdir(exist_ok=True)
                
            # Write settings to file
            with open(settings_file, 'w') as f:
                json.dump(settings, f, indent=2)
                
        except Exception as e:
            # Log error but don't crash the application
            print(f"DEBUG [APP]: Error saving color settings: {e}")

def convert_color_to_hex(color) -> str:
    """
    Convert any color format (RGB tuple, RGBA tuple, or string) to hex format for tkinter.
    
    Args:
        color: Color in any format supported by matplotlib (RGB tuple, RGBA tuple, hex string, etc.)
        
    Returns:
        Color in hex string format (#RRGGBB)
    """
    import matplotlib.colors as mcolors
    
    # If it's already a hex color, return it
    if isinstance(color, str) and color.startswith('#') and len(color) in (7, 9):
        return color[:7]  # Strip alpha if present
    
    # Convert to RGB if it's a string representation of RGB or RGBA values
    if isinstance(color, str):
        try:
            # Try to interpret as a matplotlib color
            rgba = mcolors.to_rgba(color)
            return '#{:02x}{:02x}{:02x}'.format(
                int(rgba[0] * 255),
                int(rgba[1] * 255),
                int(rgba[2] * 255)
            )
        except ValueError:
            # Default to a safe color if conversion fails
            return "#1f2937"
    
    # If it's a tuple of RGB or RGBA values
    if isinstance(color, tuple) or isinstance(color, list):
        # Ensure values are in 0-1 range
        rgb = color[:3]  # Use just the RGB part
        return '#{:02x}{:02x}{:02x}'.format(
            int(rgb[0] * 255),
            int(rgb[1] * 255),
            int(rgb[2] * 255)
        )
    
    # Default color if all else fails
    return "#1f2937"

def setup_app_icon(root: tk.Tk) -> None:
    """
    Set up the application icon for the main window and all dialogs.
    
    This replaces the default feather icon with a custom icon from the assets folder.
    
    Args:
        root: The main Tkinter root window
    """
    # Try various icon formats and locations
    icon_paths = []
    
    if platform.system() == "Windows":
        # Windows: Try .ico first, fall back to .png
        icon_paths = [
            Path("assets/icon.ico"),
            Path("assets/gantt_icon.ico"),
            Path("assets/icon.png"),
            Path("assets/gantt_icon.png")
        ]
    else:
        # macOS/Linux: Try .png first, then .ico
        icon_paths = [
            Path("assets/icon.png"),
            Path("assets/gantt_icon.png"),
            Path("assets/icon.ico"),
            Path("assets/gantt_icon.ico")
        ]
    
    # Try each icon path until one works
    for icon_path in icon_paths:
        if not icon_path.exists():
            debug_print(f"Icon file not found: {icon_path}")
            continue
            
        try:
            if platform.system() == "Windows":
                try:
                    # Try setting the window icon using iconbitmap first (Windows preferred)
                    root.iconbitmap(default=str(icon_path))
                    debug_print(f"Set application icon with iconbitmap: {icon_path}")
                    return
                except tk.TclError:
                    # If .ico format fails, try PhotoImage method instead
                    if str(icon_path).lower().endswith('.png'):
                        img = tk.PhotoImage(file=str(icon_path))
                        root.iconphoto(True, img)
                        debug_print(f"Set application icon with iconphoto: {icon_path}")
                        return
            else:
                # For Linux and macOS
                img = tk.PhotoImage(file=str(icon_path))
                root.iconphoto(True, img)
                debug_print(f"Set application icon with iconphoto: {icon_path}")
                return
                
        except Exception as e:
            debug_print(f"Error setting application icon with {icon_path}: {e}")
    
    debug_print("Could not set any application icon")

# Modify the __main__ section to destroy the unwanted tk window

if __name__ == "__main__":
    debug_print("In __main__")
    use_system_appearance()
    
    # Find all existing Tk windows before creating our own
    debug_print("Looking for existing Tk instances")
    import gc
    tk_windows = []
    
    # First check if tk._default_root already exists (created by another module)
    if tk._default_root is not None:
        debug_print(f"Found existing tk._default_root: {tk._default_root}, title: {tk._default_root.title()}")
        if tk._default_root.title() == "tk":
            debug_print("Destroying default root window with title 'tk'")
            tk._default_root.withdraw()  # Hide it
            tk._default_root.destroy()
            tk._default_root = None
    
    # Now create our window
    debug_print("Creating our root Tk instance")
    root = tk.Tk()
    root.title("QuickGantt")  # Set title immediately
    
    # Run a final check after a slight delay to catch any other windows
    def finalize_setup():
        debug_print("Finalizing setup and checking for extra windows")
        extra_windows = []
        for obj in gc.get_objects():
            if isinstance(obj, tk.Tk) and obj != root:
                extra_windows.append(obj)
                debug_print(f"Found extra Tk window: {obj}, title: {obj.title()}")
                if obj.title() == "tk":
                    debug_print("Attempting to destroy extra 'tk' window")
                    try:
                        obj.withdraw()
                        obj.destroy()
                        debug_print("Extra 'tk' window destroyed")
                    except Exception as e:
                        debug_print(f"Error destroying extra window: {e}")
        
        if not extra_windows:
            debug_print("No extra windows found, app is clean!")
    
    # Run our cleanup after a brief delay
    root.after(100, finalize_setup)
    
    # Make sure tk._default_root is set to our window
    tk._default_root = root
    
    # Set up the application icon
    setup_app_icon(root)
    
    # Continue with app creation
    debug_print("Creating QuickGanttApp")
    app = QuickGanttApp(root)
    debug_print("Starting mainloop")
    root.mainloop()