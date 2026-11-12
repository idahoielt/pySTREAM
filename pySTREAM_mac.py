'''pySTREAM 2.0 Energy scenario modelling tool, created by Ida Cecilie Hoielt as a Bachelors thesis project at DTU 2023.
Based on STREAM modelling tool created in collaboration with DTU
'''

import os
import sys
import threading
import tkinter as tk
from tkinter import *
from tkinter import ttk
import tkinter.font as font
import pandas as pd
import openpyxl
from PIL import Image, ImageTk
import plotly.io as pio
pio.renderers.default='browser'
import matplotlib
#Plots are drawn with Tk, like the rest of the window. On macOS, matplotlib would otherwise use Apple's own
#window system ('macosx'), and two window systems in one program can freeze or crash it. (Windows already uses Tk.)
matplotlib.use('TkAgg')
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import plotly.graph_objects as go
import ctypes
import tkinter.messagebox
from tkinter import filedialog
import shutil
import traceback
from datetime import datetime

#Errors are shown in a window in plain language and the details are saved in a log file,
#instead of only being printed in the terminal (which students often do not see)
ERROR_LOG = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'pySTREAM error log.txt')

def friendly_error_message(error):
    """Explain an error in plain language."""
    if isinstance(error, PermissionError):
        return ("pySTREAM could not open or save this file:\n" + str(error.filename) +
                "\n\nThe file is probably open in Excel or another program. Close it there and try again.")
    if isinstance(error, FileNotFoundError):
        return ("pySTREAM could not find this file:\n" + str(error.filename) +
                "\n\nCheck that it is in the pySTREAM folder, or download pySTREAM again.")
    if isinstance(error, KeyError):
        return ("A value pySTREAM needs is missing in the data files: " + str(error) +
                "\n\nCheck that the data files are complete and that no names in them were changed.")
    return "Something went wrong: " + type(error).__name__ + ": " + str(error)

def handle_error(error_type, error, error_traceback):
    """Show an error in a window and save the details in the error log."""
    details = ''.join(traceback.format_exception(error_type, error, error_traceback))
    print(details, file=sys.stderr)
    message = friendly_error_message(error)
    try:
        with open(ERROR_LOG, 'a', encoding='utf-8') as log:
            log.write(datetime.now().strftime('--- %Y-%m-%d %H:%M:%S\n') + details + '\n')
        message += "\n\nThe details are saved in the error log:\n" + ERROR_LOG + "\nPlease include that file if you ask for help."
    except OSError:
        pass
    tk.messagebox.showerror("pySTREAM - something went wrong", message)

sys.excepthook = handle_error #errors while pySTREAM starts (errors after a button is pressed: see window.report_callback_exception)

print('Starting pySTREAM...')
from Backend import *
from Time_Series import run_timeSeries


#This class is a copy of the tkinter Treeview class that is used to create tables. 
#This class adds the ability to edit cells directly in the table
#allowed: the lowest and highest value a cell may get (None = no maximum); what and unit_label are used in the message
#that explains it, e.g. "Shares must be between 0 % and 100 %."
class TreeviewEdit(ttk.Treeview):
    def __init__(self, master, parent_frame, y_pos, x_pos, unit, allowed=(0, 100), what='Shares', unit_label='%', **kwargs):
        super().__init__(master, **kwargs)

        self.parent_frame = parent_frame
        self.y_pos = y_pos
        self.x_pos = x_pos
        self.unit = unit
        self.allowed = allowed
        self.what = what
        self.unit_label = unit_label

        self.bind("<Double-1>", self.on_double_click) #enter cell edit when a cell is double clicked
    
    #Creates tk.Entry widget on top of the selected cell to enter new data there
    def on_double_click(self, event):
        #event has the coordinates of the mouse when double click occurred
        
        #Identify the region that is double clicked
        region_clicked = self.identify_region(event.x, event.y)
       
        #Only continue if the clicked region is a cell in the table
        if region_clicked not in "cell":
            return
        
        #Identify which cell was double clicked, for example #1 is the first column, #2 is the second, etc.
        column = self.identify_column(event.x)
        
        column_index = int(column[1:]) - 1
        #Only continue if the first column was not clicked
        if column in "#1":
            return
        
        #item iid, for example 001
        selected_iid = self.focus()
        #Rows that name a group of rows (Person, Goods, ...) can not be edited
        if GROUP[0] in self.item(selected_iid, 'tags'):
            return
        #Contains values from the row from the given item iid
        selected_values = self.item(selected_iid)
        #Only the text from the column that the cell clicked is in
        selected_text = selected_values.get('values')[column_index]
        column_box = self.bbox(selected_iid, column)
        
        #Create entry widget that covers the cell that was clicked
        entry_edit = ttk.Entry(self.parent_frame, font=font_table) #the same text size as the cell, on every screen
        
        #Record the column index and item iid
        entry_edit.editing_column_index = column_index
        entry_edit.editing_item_iid = selected_iid
        entry_edit.insert(0, selected_text)
        entry_edit.select_range(0, tk.END) #select the text in the cell
        
        
        entry_edit.focus()
        
        entry_edit.bind("<FocusOut>", self.on_focus_out) #on click away from cell when editing cell
        entry_edit.bind("<Return>", self.on_enter_pressed) #on enter press on keyboard when editing cell
        
        entry_edit.place(x=column_box[0]+self.x_pos, y=column_box[1]+self.y_pos, w=column_box[2], h=column_box[3])
    
    #When user is editing cell in table and hits enter on keyboard it will save the entered data to the table
    def on_enter_pressed(self, event):
        new_value = event.widget.get()
        
        #for example I002
        selected_iid = event.widget.editing_item_iid
        
        #for example 1 (second column, first column with values)
        column_index = event.widget.editing_column_index
        
        event.widget.destroy()
        
        #Read the number; a % sign and a decimal comma (5,3) are accepted
        try:
            number = float(new_value.strip().rstrip('%').strip().replace(',', '.'))
        except ValueError:
            tk.messagebox.showerror("Invalid input", "Please enter a number, for example 5.3 (in " + self.unit_label + ").")
            return
        
        #Check that the number is allowed, and explain the allowed values if not (the table keeps its old value)
        low, high = self.allowed
        if number < low or (high is not None and number > high):
            if high is None:
                rule = self.what + " must be " + format(low, 'g') + " " + self.unit_label + " or more."
            else:
                rule = self.what + " must be between " + format(low, 'g') + " " + self.unit_label + " and " + format(high, 'g') + " " + self.unit_label + "."
            tk.messagebox.showerror("Invalid input", rule + "\n\nYou entered: " + new_value.strip())
            return
        
        if self.unit == '%': # If the unit in the table is a percentage (%), e.g. 5.3 is shown as 5.30%
            formatted_value = f'{number / 100:.2%}'
        else:
            formatted_value = f'{number:.2f}'
        current_values = self.item(selected_iid).get("values")
        current_values[column_index] = formatted_value
        self.item(selected_iid, values=current_values)
    
    #When user clicks out of editing cell without saving by pressing enter the new value entered is forgotten
    def on_focus_out(self, event):
        event.widget.destroy()

class FlatButton(tk.Label):
    """A button that keeps its colours on every system. On macOS, tk.Button ignores the background colour and
    is drawn as a light grey Mac button, so the white text of most pySTREAM buttons could not be read.
    It is used like tk.Button: command runs when the button is clicked (or invoke() is called), and
    activebackground/activeforeground are the colours while the mouse button is held down."""
    def __init__(self, master=None, command=None, **options):
        #the space around the text that a Windows button has (it matters for buttons sized by their text)
        options.setdefault('padx', px(3))
        options.setdefault('pady', px(4))
        super().__init__(master, **options)
        self.command = command
        self.bind('<ButtonPress-1>', lambda event: self.config(state='active'))
        self.bind('<ButtonRelease-1>', self.on_release)

    def on_release(self, event):
        self.config(state='normal')
        #like a real button, letting go of the mouse outside the button does not click it
        if 0 <= event.x < self.winfo_width() and 0 <= event.y < self.winfo_height():
            self.invoke()

    def invoke(self):
        if self.command is None:
            return None
        try:
            return self.command()
        except Exception:
            #as with tk.Button, errors are shown in pySTREAM's error window (see handle_error)
            self._root().report_callback_exception(*sys.exc_info())

GROUP = ('group',) #tag of the rows that name the rows below them (Person, Goods, ...)

def show_group_rows(table):
    """Rows tagged GROUP look like sub-headings: light blue and bold. TreeviewEdit does not let them be edited."""
    table.tag_configure(GROUP[0], background=c_lightish, foreground=c_dark, font=font_table_heading)

def side_table(frame, heading, width, style):
    """Read-only table with one column, placed next to another table (amounts, sums).
    It has the same heading and row height as the table next to it, so their rows line up on every screen."""
    table = ttk.Treeview(frame, columns=['value'], show='headings', selectmode='none', style=style)
    table.heading('value', text=heading, anchor='center')
    table.column('value', width=width, anchor='center', stretch=False)
    return table

#Used to save user input in tables to the corresponding dataframe when the corresponding save button is pressed
def save_table_to_dataframe(treeview, has_children):
    #Get the column headings
    columns = [treeview.heading(column_id)["text"] for column_id in treeview["columns"]]

    #Create an empty list to store data
    data = []

    # Iterate through the Treeview items and append data to the list
    for item in treeview.get_children():
        
        if has_children: #if the table has parents with children, only the childrens values will be stored in dataframe
            children = treeview.get_children(item)
            for child in children:
                values = treeview.item(child, 'values')
                data.append(values)
            if treeview.parent(item) == '':
                continue  # Skip parent items
            
        values = treeview.item(item, "values")
        data.append(values)
        
    
    # Create a DataFrame from the list of data
    df = pd.DataFrame(data, columns=columns)
    df.set_index(columns[0], inplace=True) #Using the first table column as index in the dataframe
    
    return df

#Global constants
SECTORS = ['Residential', 'Tertiary', 'Industry', 'Agriculture']

#Plots are also saved as image files in the 'figures' folder, so they can be used in reports
FIGURES_DIR = os.path.join(PROGRAM_DIR, 'figures')
os.makedirs(FIGURES_DIR, exist_ok=True)

def figure_path(file_name):
    return os.path.join(FIGURES_DIR, file_name)

#colors
c_white = '#F5F5F5'
c_black = '#2D2E2E'
c_dark = '#27567C'
c_light = '#D2E1EF'
c_lightish = '#B2CEE4'
c_mid = '#7CACD1'
c_midlight = '#92BAD8'
c_accent = '#F1D89C'


#Open program window
#==============================================================================
#The layout is designed for a screen of 1920 x 1080 pixels, but the pages only use the top left USED_WIDTH x USED_HEIGHT
#pixels of it; the rest is empty. When pySTREAM starts, the layout is scaled by SCALE so that this used part fills the
#screen it is shown on. The layout is then ww x wh pixels: positions and sizes are fractions of ww and wh, text is scaled with
#'tk scaling' (below), and the few sizes given in pixels use px(). On very small screens the layout keeps a readable
#minimum size and scrollbars appear. The layout sits in a scrollable area of the window (root).
DESIGN_WIDTH = 1920
DESIGN_HEIGHT = 1080
USED_WIDTH = 1880  #the pages end at about 1862 pixels
USED_HEIGHT = 1010 #the lowest button in the menu ends at about 993 pixels
MIN_SCALE = 0.6 #smaller text is hard to read

#On Windows with display scaling above 100 %, Windows would otherwise enlarge the whole window as a picture (blurry
#text, and the layout did not fit the screen), and the first plot made matplotlib change this setting halfway through.
if sys.platform == 'win32':
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1) #pySTREAM scales itself (see SCALE)
    except (AttributeError, OSError):
        pass #older Windows: Windows scales the window as before

root = tk.Tk()
root.report_callback_exception = handle_error #errors after a button is pressed
root.title("pySTREAM - Energy System Simulation Tool")
root.geometry('%dx%d+0+0' % (min(DESIGN_WIDTH, root.winfo_screenwidth()), min(DESIGN_HEIGHT, root.winfo_screenheight())))
try:
    root.state('zoomed') #fill the screen (Windows and macOS)
except tk.TclError:
    pass #not available on every system; the size set above is used instead

def available_size():
    """Width and height pySTREAM can use: the inside of the window when it fills the screen."""
    root.update()
    if root.state() == 'zoomed' and root.winfo_width() > 200 and root.winfo_height() > 200:
        return root.winfo_width(), root.winfo_height()
    return root.winfo_screenwidth(), root.winfo_screenheight() - 100 #leave room for the taskbar and title bar

def layout_scale(width, height):
    """How much the 1920 x 1080 layout is enlarged (> 1) or reduced (< 1) so its used part fits a space of width x height pixels."""
    return max(MIN_SCALE, min(width / USED_WIDTH, height / USED_HEIGHT))

SCALE = layout_scale(*available_size())
ww = int(DESIGN_WIDTH * SCALE)  #layout width
wh = int(DESIGN_HEIGHT * SCALE) #layout height
used_width = round(USED_WIDTH * SCALE)   #the part of the layout that has content
used_height = round(USED_HEIGHT * SCALE)
if root.state() != 'zoomed': #the window could not be made to fill the screen: give it the size of the used part of the layout
    root.geometry('%dx%d+0+0' % (min(used_width, root.winfo_screenwidth()), min(used_height, root.winfo_screenheight() - 100)))

def px(pixels):
    """A size in pixels of the 1920 x 1080 design, scaled to this screen."""
    return max(1, round(pixels * SCALE))

#Text sizes are given in points (with pt() below). The design uses 96 pixels per inch (72 points), the Windows
#standard. On Windows and Linux, 'tk scaling' scales all text in the same way as the layout. Tk on macOS ignores
#'tk scaling' (one point is always one screen point there), so on macOS pt() scales the sizes themselves instead.
if sys.platform == 'darwin':
    FONT_SCALE = SCALE * 96 / 72
else:
    FONT_SCALE = 1
    root.tk.call('tk', 'scaling', SCALE * 96 / 72)

def pt(size):
    """A text size in points of the 1920 x 1080 design, for this screen."""
    return max(1, round(size * FONT_SCALE))

#Plots in the window are scaled in the same way: on Windows and Linux matplotlib does this itself (it follows
#'tk scaling'), on macOS it is done here. The matplotlib plots saved in the figures folder keep the same size on every screen.
if sys.platform == 'darwin':
    plt.rcParams['figure.dpi'] = 100 * SCALE
plt.rcParams['savefig.dpi'] = 100

scroll_canvas = tk.Canvas(root, bg=c_white, highlightthickness=0, scrollregion=(0, 0, used_width, used_height), xscrollincrement=40, yscrollincrement=40)
scroll_x = ttk.Scrollbar(root, orient='horizontal', command=scroll_canvas.xview)
scroll_y = ttk.Scrollbar(root, orient='vertical', command=scroll_canvas.yview)
scroll_canvas.configure(xscrollcommand=scroll_x.set, yscrollcommand=scroll_y.set)
scroll_canvas.grid(row=0, column=0, sticky='nsew')
root.grid_rowconfigure(0, weight=1)
root.grid_columnconfigure(0, weight=1)

#Everything in pySTREAM is placed in this frame of ww x wh pixels
window = tk.Frame(scroll_canvas, width=ww, height=wh)
scroll_canvas.create_window(0, 0, window=window, anchor='nw')

def show_scrollbars_when_needed(event):
    """Show the scrollbars only when the window is smaller than the used part of the layout."""
    if event.widget is not root:
        return
    if root.winfo_width() < used_width:
        scroll_x.grid(row=1, column=0, sticky='ew')
    else:
        scroll_x.grid_remove()
    if root.winfo_height() < used_height:
        scroll_y.grid(row=0, column=1, sticky='ns')
    else:
        scroll_y.grid_remove()
root.bind('<Configure>', show_scrollbars_when_needed)

def scroll_with_mouse_wheel(event):
    """Scroll with the mouse wheel (Shift + wheel: sideways), except over tables, which scroll themselves."""
    if isinstance(event.widget, (ttk.Treeview, str)):
        return
    steps = -3 if (event.num == 4 or event.delta > 0) else 3
    if event.state & 0x0001: #Shift is held down
        scroll_canvas.xview_scroll(steps, 'units')
    else:
        scroll_canvas.yview_scroll(steps, 'units')
for wheel_event in ('<MouseWheel>', '<Shift-MouseWheel>', '<Button-4>', '<Button-5>'): #Button-4/5: the mouse wheel on Linux
    root.bind_all(wheel_event, scroll_with_mouse_wheel)

#Stop program when window is closed
def quit_me():
    print('Closing pySTREAM')
    plt.close()
    root.quit()
    root.destroy()
    #Stop right away: background threads that Kaleido starts for the Sankey images
    #can otherwise keep Python (and the terminal window) running after the window is closed
    sys.stdout.flush()
    os._exit(0)
root.protocol("WM_DELETE_WINDOW", quit_me)

window.configure(bg=c_white)

#Show that pySTREAM is starting, until the window is ready (see the end of this file)
starting_label = tk.Label(root, text="Starting pySTREAM...\nThis can take up to a minute.", font=('Arial', pt(16)), bg=c_white, fg=c_dark)
starting_label.place(relx=0.5, rely=0.5, anchor='center')
root.update()

#Long calculations show a message and a busy cursor while they run (the window cannot react in the meantime)
busy_label = tk.Label(root, bg=c_accent, fg=c_dark, font=('Arial', pt(14), 'bold'))
def run_while_busy(action, message):
    busy_label.config(text=message)
    busy_label.place(relx=0.5, rely=0.5, anchor='center', width=ww*0.3, height=wh*0.06)
    busy_label.lift()
    root.config(cursor='watch')
    root.update()
    try:
        action()
    finally:
        busy_label.place_forget()
        root.config(cursor='')

#fonts
font_header = font.Font(family='Nueva Std Cond', size=pt(24), weight='bold')
font_subheader= font.Font(family='Nueva Std Cond', size=pt(20), weight='bold')
font_subsubheader= font.Font(family='Nueva Std Cond', size=pt(17), weight='bold')
font_subsubsubheader = font.Font(family='Nueva Std Cond', size=pt(14), weight='bold')
font_button = font.Font(family='Nueva Std Cond', size=pt(12), weight='bold')
font_text = font.Font(family='Nueva Std Cond', size=pt(10))
font_text_bold = font.Font(family='Nueva Std Cond', size=pt(10), weight='bold')
font_table_heading = ('Arial', pt(9), 'bold')
font_table = ('Arial', pt(9))

#table style
style = ttk.Style()
style.theme_use('default')
style.configure('Treeview', background = c_white, foreground = c_black, rowheight = px(23), fieldbackground = c_white, relief = 'flat', borderwidth = 0, font = font_table)  
style.map('Treeview', background=[('selected', c_midlight)], foreground=[('selected', c_black)])
style.layout('Treeview', [('Treeview.treearea', {'sticky': 'nswe'})]) #tables without a border (it showed as a white outline around the headings)
style.configure('Amounts.Treeview', background = c_lightish, fieldbackground = c_lightish, font = ('Arial', pt(9), 'italic')) #read-only amounts next to the production shares
style.configure('Sums.Treeview', font = ('Arial', pt(9), 'italic')) #read-only sums next to the transport fuel shares
style.configure('Treeview.Heading', background = c_mid, foreground = c_white, relief = 'flat', font=font_table_heading, padding = (px(3), 0)) #heading text starts where the cell text below it starts
style.map('Treeview.Heading', background = [('active', c_mid)])
style.configure('Treeview.Cell', padding = (px(4), 0)) #space between the left edge of a cell and its text, as in the headings
style.configure('Treeview.Item', padding = 0)
style.configure('TEntry', background = c_accent, foreground = c_black, font = font_table, borderwidth = 0, relief = 'flat', fieldbackground = c_accent, selectbackground = c_dark, selectborderwidth = 0, selectforeground = c_white)
style.map('TEntry', selectBackground = [('active', c_midlight)])
style.configure("Vertical.TScrollbar", background=c_mid, borderwidth = 0, arrowcolor=c_light, troughcolor = c_light, relief = 'flat')
style.map('Vertical.TScrollbar', background=[('pressed', '!disabled', c_mid), ('active', c_mid)])


#Creating menu on left side with buttons for navigation
#==============================================================================
menu_frame = tk.Frame(window, width=ww*0.15, height=wh, bg=c_mid)
menu_frame.place(x=0, y=0)
#"py" and "STREAM" side by side without padding, so they touch whatever font the computer uses
logo_frame = tk.Frame(menu_frame, bg=c_mid)
logo_frame.place(x=ww*0.016, y=wh*0.023)
menu_frame_label1 = tk.Label(logo_frame, text="py", font=font_header, bg=c_mid, fg=c_dark, bd=0, padx=0)
menu_frame_label1.pack(side='left')
menu_frame_label2 = tk.Label(logo_frame, text="STREAM", font=font_header, bg=c_mid, fg=c_black, bd=0, padx=0)
menu_frame_label2.pack(side='left')

button_width = ww*0.12
button_height = wh*0.07

#Action performed when graphs button is clicked
def graphs_button_action():
    select_frame("graphs_frame")

#graps button setup
graphs_button = FlatButton(menu_frame, text="Graphs", command=graphs_button_action, bd=0, bg=c_midlight, fg=c_white, activebackground=c_dark, activeforeground=c_white, font=font_button)
graphs_button.place(x=ww*0.014, y=wh*0.08*1, width=button_width, height=button_height)
#graphs_button.config(width=button_width, height=button_height, font=font_button)

#Action performed when consumption button is clicked
def consumption_button_action():
    select_frame("consumption_frame")
    
#consumption button setup
consumption_button = FlatButton(menu_frame, text="Edit consumption", command=consumption_button_action, bd=0, bg=c_midlight, fg=c_white, activebackground=c_dark, activeforeground=c_white, font=font_button)
consumption_button.place(x=ww*0.014, y=wh*0.08*2, width=button_width, height=button_height)
#consumption_button.config(width=button_width, height=button_height, font=font_button)

#Action performed when production button is clicked
def production_button_action():
    select_frame("production_frame")

#production button setup
production_button = FlatButton(menu_frame, text="Edit production", command=production_button_action, bd=0, bg=c_midlight, fg=c_white, activebackground=c_dark, activeforeground=c_white, font=font_button)
production_button.place(x=ww*0.014, y=wh*0.08*3, width=button_width, height=button_height)
#production_button.config(width=button_width, height=button_height, font=font_button)

#Action performed when base year button is clicked
def by_button_action():
    select_frame("by_frame")

#base year button setup
by_button = FlatButton(menu_frame, text="Base year data", command=by_button_action, bd=0, bg=c_midlight, fg=c_white, activebackground=c_dark, activeforeground=c_white, font=font_button)
by_button.place(x=ww*0.014, y=wh*0.08*4, width=button_width, height=button_height)
#by_button.config(width=button_width, height=button_height, font=font_button)  
    
#Action performed when scenario button is clicked
def scenario_button_action():
    select_frame("scenario_frame")

#scenario button setup
scenario_button = FlatButton(menu_frame, text="Scenario data", command=scenario_button_action, bd=0, bg=c_midlight, fg=c_white, activebackground=c_dark, activeforeground=c_white, font=font_button)
scenario_button.place(x=ww*0.014, y=wh*0.08*5, width=button_width, height=button_height)
#scenario_button.config(width=button_width, height=button_height, font=font_button)

#Action performed when transport button is clicked
def transport_button_action():
    select_frame("transport_frame")
    
#transport button setup
transport_button=FlatButton(menu_frame, text="Transport data", command=transport_button_action, bd=0, bg=c_midlight, fg=c_white, activebackground=c_dark, activeforeground=c_white, font=font_button)
transport_button.place(x=ww*0.014, y=wh*0.08*6, width=button_width, height=button_height)
#transport_button.config(width=button_width, height=button_height, font=font_button)

#Action performed when CO2 button is clicked
def CO2_button_action():
    select_frame("CO2_frame")
    
#CO2 button setup
CO2_button = FlatButton(menu_frame, text="CO2 emissions", command=CO2_button_action, bd=0, bg=c_midlight, fg=c_white, activebackground=c_dark, activeforeground=c_white, font=font_button)
CO2_button.place(x=ww*0.014, y=wh*0.08*7, width=button_width, height=button_height)
#CO2_button.config(width=button_width, height=button_height, font=font_button)


#Action performed when economy button is clicked
def economy_button_action():
    select_frame("economy_frame")
    
#Economy button setup
economy_button = FlatButton(menu_frame, text="Economy", command=economy_button_action, bd=0, bg=c_midlight, fg=c_white, activebackground=c_dark, activeforeground=c_white, font=font_button)
economy_button.place(x=ww*0.014, y=wh*0.08*8, width=button_width, height=button_height)
#economy_button.config(width=button_width, height=button_height, font=font_button)

#Action performed when time series button is clicked
def timeSeries_button_action():
    select_frame("timeSeries_frame")

#Time series button setup
timeSeries_button = FlatButton(menu_frame, text="Time series", command=timeSeries_button_action, bd=0, bg=c_midlight, fg=c_white, activebackground=c_dark, activeforeground=c_white, font=font_button)
timeSeries_button.place(x=ww*0.014, y=wh*0.08*9, width=button_width, height=button_height)
#timeSeries_button.config(width=button_width, height=button_height, font=font_button)

#==============================================================================


#Change to the selected frame and forget all others
#==============================================================================    
def select_frame(frame):
    graphs_frame.place_forget()
    consumption_frame.place_forget()
    production_frame.place_forget()
    by_frame.place_forget()
    scenario_frame.place_forget()
    transport_frame.place_forget()
    CO2_frame.place_forget()
    economy_frame.place_forget()
    timeSeries_frame.place_forget()
    
    if frame == "graphs_frame":
        show_graphs_frame()

    elif frame == "consumption_frame":
        show_consumption_frame()
        
    elif frame == "production_frame":
        show_production_frame()
        
    elif frame == "by_frame":
        show_by_frame()
    
    elif frame == "scenario_frame":
        show_scenario_frame()
   
    elif frame == "transport_frame":
        show_transport_frame()
    
    elif frame == "CO2_frame":
        show_CO2_frame()
        
    elif frame == "economy_frame":
        show_economy_frame()

    elif frame == "timeSeries_frame":
        show_timeSeries_frame()

#==============================================================================

#Functions to create multiple similar plots
#==============================================================================

PLOT_COLORS = ['#FEAA00', '#E467AC', '#A96CCE', '#4266BC', '#83D2F4', '#FFFF32', '#217378', '#AF5500', '#dc2f02', '#C5F567', '#5e60ce']

#Applied to all plots to make them look better :)
def plot_standards(fig, ax):
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(False)
    ax.spines['bottom'].set_color('#DDDDDD')
    ax.tick_params(bottom=False, left=False)
    ax.set_axisbelow(True)
    
#Stacked bar plot 
def plot_stacked_bar(data, legend, columns, figsize):
    fig, ax = plt.subplots(figsize=figsize)
    data.iloc[:, :columns].T.plot(kind="bar", stacked=True, ax=ax, color=(PLOT_COLORS))
    plot_standards(fig, ax)
    ax.yaxis.grid(True, color='#EEEEEE')
    ax.legend(loc='center left', bbox_to_anchor=(1, 0.5), title=legend, alignment='left', facecolor=c_light, framealpha=0)
    ax.tick_params(axis='x', labelrotation=0)
    ax.set_facecolor(c_lightish)
    fig.set_facecolor(c_light)
    plt.tight_layout(pad=2)
    return fig, ax

#Bar plot 
def plot_bar(data):
    fig, ax = plt.subplots(figsize=(8, 6))
    data.plot(kind="bar", stacked=False, ax=ax, color=(PLOT_COLORS))
    plot_standards(fig, ax)
    ax.yaxis.grid(True, color='#EEEEEE')
    ax.set_ylabel("[PJ]")
    ax.tick_params(axis='x', labelrotation=50)
    ax.set_facecolor(c_lightish)
    fig.set_facecolor(c_light)
    plt.tight_layout(pad=3.7) 
    return fig, ax

#Horizontal bar plot
def plot_horizontal_bar(data):
    fig, ax = plt.subplots(figsize=(8, 6))
    data.plot(kind="barh", stacked=False, ax=ax, color=(PLOT_COLORS))
    plot_standards(fig, ax)
    ax.xaxis.grid(True, color='#EEEEEE')  
    ax.set_xlabel("[PJ]")  
    ax.tick_params(axis='y', labelrotation=0)  
    ax.set_facecolor(c_lightish)
    fig.set_facecolor(c_light)
    plt.tight_layout(pad=3.7) 
    return fig, ax

#Pie plot
def plot_pie(data, colors):
    labels = data.index
    fig, ax = plt.subplots(figsize=(8, 8))
    patches, texts, autotexts = ax.pie(data, startangle=90, labels=labels, autopct='%1.1f%%', colors=colors, labeldistance=0.7, pctdistance=0.3, wedgeprops=dict(alpha=0.6))
    for text in texts: #Change 
        text.set_horizontalalignment('center')
        text.set(color=c_black, fontsize=7)
        text.set_fontstyle('italic')
        
    for autotext in autotexts:
        autotext.set_horizontalalignment('center')
        autotext.set(color=c_white, weight='bold', fontsize=9)
        
    ax.axis('equal')  
    ax.set_facecolor(c_lightish)
    fig.set_facecolor(c_light)
    
    plt.tight_layout(pad=0.2)

    return fig, ax

#Duration curve plot
def plot_duration_curve(duration_curve, technologies_used):
    bar_height = technologies_used['Avg'].tolist()
    bar_width = technologies_used['Hours'].tolist()
    bar_legend = technologies_used.index.tolist()

    # Create the plot
    fig, ax = plt.subplots(figsize=(8, 6))

    # Plot the duration curve
    ax.plot(duration_curve, label='Duration Curve', color=c_black, linewidth=2)

    # Plot the horizontal bars on top
    ypos = []
    a = 0
    for i in range(len(bar_legend)):
        if i == 0:
            a+=bar_height[i]
            ypos.append(bar_height[i]/2)
        else:
            a += bar_height[i]
            ypos.append(a - bar_height[i]/2)

        ax.barh(y=ypos[i], width=bar_width[i], height=bar_height[i], color=PLOT_COLORS[i], edgecolor=PLOT_COLORS[i], label=bar_legend[i], alpha=0.7)

    # Customize the plot¨9
    ax.spines['bottom'].set_position('zero')
    ax.xaxis.set_ticks_position('bottom')
    ax.xaxis.set_tick_params(labelbottom=True, labeltop=False)

    ax.set_xlabel('Hours', fontsize=10, fontweight='bold')
    ax.set_ylabel('Avg. Power [MW]', fontsize=8)
    #y-axis from the lowest to the highest value of the curve, rounded out to whole 1000 MW
    ax.set_ylim(np.floor(min(duration_curve.min(), 0) / 1000) * 1000, np.ceil(max(duration_curve.max(), 0) / 1000) * 1000)
    ax.set_title('Load Duration Curve')
    ax.set_facecolor(c_lightish)
    ax.yaxis.grid(True, color='#EEEEEE') 
    ax.set_axisbelow(True) 
    ax.axhline(y=0, color='black', linewidth=1.5, alpha=1) 
    ax.legend(loc='upper right', alignment='left', facecolor=c_light, framealpha=0.5, fontsize=7, title='Flexible Technologies',title_fontsize=8)

    fig.set_facecolor(c_light)

    # Apply plot standards
    plot_standards(fig, ax)

    return fig, ax

timeSeries_colors = ['#C5F567', '#A96CCE', '#FF7400',  '#5e60ce','#83D2F4', '#38ACAE', '#FEAA00', '#66E067', '#E467AC', '#FFFF32', '#4266BC', '#217378', '#AF5500', '#808080', '#f94144', '#f3722c', '#f8961e']

#Time series average week plot 
def plot_timeSeries_avgWeek(df_hour):

    df_hour['Exported electricity'] = -abs(df_hour['Exported electricity'])
    # Calculate groupings
    df_hour['HourOfWeek'] = df_hour.index % 168

    # Separate out the data for positive and negative values
    positive_technologies = df_Power_SC.index
    negative_technologies = ['Exported electricity']

    # Aggregate data for average week and average month
    avg_week = df_hour.groupby('HourOfWeek')[positive_technologies].mean()
    avg_week_exp = df_hour.groupby('HourOfWeek')[negative_technologies].mean()
    fig, ax = plt.subplots(figsize=(20, 5))

    #plot the average production of each technology for consumption
    ax.stackplot(avg_week.index, avg_week[positive_technologies].T, labels=positive_technologies, colors=timeSeries_colors, alpha=0.7)
    
    #plot the average export as negative values
    ax.stackplot(avg_week_exp.index, avg_week_exp[negative_technologies].T, labels=negative_technologies, colors=['#FF7400'], alpha=0.7)

    ax.set_xticks(np.arange(12, 169, 24))
    ax.set_xticklabels(['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'])
    ax.set_ylabel('Power [MW]', fontsize=8)
    ax.set_title('Average Weekly Power Production by Technology')
    ax.set_facecolor(c_lightish)
    ax.set_xlim(0, 167)
    fig.set_facecolor(c_light)
    fig.set_snap(True)
    ax.yaxis.grid(True, color='#EEEEEE') 
    ax.legend(loc='center left', bbox_to_anchor=(1, 0.5), alignment='left', facecolor=c_light, framealpha=0, fontsize=7)
    plot_standards(fig, ax)
    
    return fig, ax

#Time series average month plot
def plot_timeSeries_avgMonth(df_hour):
    df_hour['Exported electricity'] = -abs(df_hour['Exported electricity'])
    # Calculate groupings
    df_hour['HourOfMonth'] = df_hour.index % 744

    # Separate out the data for positive and negative values
    positive_technologies = df_Power_SC.index
    negative_technologies = ['Exported electricity']

    # Aggregate data for average week and average month
    avg_month = df_hour.groupby('HourOfMonth')[positive_technologies].mean()
    avg_month_exp = df_hour.groupby('HourOfMonth')[negative_technologies].mean()
    fig, ax = plt.subplots(figsize=(20, 5))

    #plot the average production of each technology for consumption
    ax.stackplot(avg_month.index, avg_month[positive_technologies].T, labels=positive_technologies, colors=timeSeries_colors, alpha=0.7)
    
    #plot the average export as negative values
    ax.stackplot(avg_month_exp.index, avg_month_exp[negative_technologies].T, labels=negative_technologies, colors=['#FF7400'], alpha=0.7)

    ax.set_xticks(np.arange(0, 745, 24*4))
    ax.set_xticklabels([f'Day {i}' for i in range(1, 32, 4)], rotation=0)
    ax.set_ylabel('Power [MW]', fontsize=8)
    ax.set_title('Average Monthly Power Production by Technology')
    ax.set_facecolor(c_lightish)
    ax.set_xlim(0, 743)
    fig.set_facecolor(c_light)
    fig.set_snap(True)
    ax.yaxis.grid(True, color='#EEEEEE') 
    ax.legend(loc='center left', bbox_to_anchor=(1, 0.5), alignment='left', facecolor=c_light, framealpha=0, fontsize=7)
    plot_standards(fig, ax)
    
    return fig, ax

#==============================================================================

#Create page headers for each page
#==============================================================================
def header_frame(frame, title, subtitle, subxpos):
    header_frame = tk.Frame(frame, width=ww*0.53, height=wh*0.07, bg=c_white)
    header_frame.place(x=0, y=0)
    header = tk.Label(header_frame, text=title, font=font_header, bg=c_white, fg=c_dark)
    header.place(x=0, y=0)
    subheader = tk.Label(header_frame, text=subtitle, font=font_subheader, bg=c_white, fg=c_mid)
    subheader.place(x=subxpos, y=px(5))
    
#==============================================================================

#Create graphs frame that appears when graphs button is pressed
#==============================================================================
graphs_frame = tk.Frame(window)
graph_width = ww*0.38
graph_height = wh*0.6
def show_graphs_frame():
    graphs_frame.config(width=ww*0.8, height=wh*0.85, bg=c_white) 
    graphs_frame.place(x=ww*0.17, y=wh*0.025)
    sankey_SC()
    sankey_BY()
    
header_frame(graphs_frame, "Graphs", "Scenario and Base year", ww*0.08)

#Base year graphs frame    
graphs_by_frame = tk.Frame(graphs_frame, bg=c_light)
graphs_by_frame.place(x=0, y=wh*0.18, height=wh*0.65, width=ww*0.39)
by_label = tk.Label(graphs_by_frame, text="Base Year", font=font_subheader, bg=c_light, fg=c_dark)
by_label.place(x=0, y=wh*0.005, width=ww*0.38)

#Scenario year graphs frame
graphs_sc_frame = tk.Frame(graphs_frame, bg=c_light)
graphs_sc_frame.place(x=ww*0.4, y=wh*0.18, height=wh*0.65, width=ww*0.39)
sc_label = tk.Label(graphs_sc_frame, text="Scenario Year", font=font_subheader, bg=c_light, fg=c_dark)
sc_label.place(x=0, y=wh*0.005, width=ww*0.38)

#Sankey energy flow diagram
#------------------------------------------------------------------------------   
def sankey_diagram(df_fuelType_consumption, df_Fc, frame):
    
    label = ['Bioenergy', 'Coal', 'Oil', 'Geothermal', 'Hydro', 'Waste', 'Natural Gas', 'Nuclear', 'Solar', 'Wave', 'Wind', #sources
           'Fuel Refinery', 'Electricity', 'District heating', #conversion 11, 12, 13
           'Transport','Residential', 'Tertiary', 'Industry', 'Agriculture'] #sectors consumption 14, 15, 16, 17, 18
                 
    #Flows from the primary energy sources (nodes 0-10) to where they are used, for every fuel type and use, so no flow
    #can be left out. The fuel for CHP heat ('Heat production') goes to the electricity node, which passes the CHP heat
    #on to district heating (see El below).
    source_targets = {'Fuel Refinery': 11, 'Electricity production': 12, 'Heat production': 12, 'DH boilers': 13,
                      'Transport': 14, 'Residential': 15, 'Tertiary': 16, 'Industry': 17, 'Agriculture': 18}
    source_colors = ['rgba(197, 245, 103, 0.65)', 'rgba(169, 108, 206, 0.65)', 'rgba(94, 96, 206, 0.65)', 'rgba(255, 116, 0, 0.65)',
                     'rgba(56, 172, 174, 0.65)', 'rgba(254, 170, 0, 0.65)', 'rgba(102, 224, 103, 0.65)', 'rgba(228, 103, 172, 0.65)',
                     'rgba(255, 255, 50, 0.65)', 'rgba(66, 102, 188, 0.65)', 'rgba(131, 210, 244, 0.65)']
    source, target, values, color_link = [], [], [], []
    for i, fuel_type in enumerate(label[:11]):
        for column, node in source_targets.items():
            value = df_fuelType_consumption.loc[fuel_type, column]
            if value > 0:
                source.append(i)
                target.append(node)
                values.append(value)
                color_link.append(source_colors[i])

    FR = df_Fc.loc[['Biodiesel', 'Biogas', 'Diesel', 'Ethanol', 'Gasoline', 'Hydrogen', 'Methanol', 'UPGR Biogas'], ['Electricity production', 'Heat production', 'DH boilers', 'Transport', 'Residential', 'Tertiary', 'Industry', 'Agriculture']].sum().tolist()
    FR_EL = df_Fc.loc['Electricity', 'Fuel Refinery']
    FR_DH = df_Fc.loc['District Heating', 'Fuel Refinery']        
    
    El = df_fuelType_consumption.loc['Electricity', ['DH boilers', 'Transport', 'Residential', 'Tertiary', 'Industry', 'Agriculture']].tolist()
    #add either consumption or production of electricity from fuel refinery 
    if FR_EL >= 0:
        El.append(FR_EL)
        FR.append(0)
    else: 
        FR.append(abs(FR_EL))
        El.append(0)
        
    heat_CHP = df_fuelType_consumption.loc['District Heating', 'Heat production']
    El[0] += abs(heat_CHP)
        
    DH = df_fuelType_consumption.loc['District Heating', ['Residential', 'Tertiary', 'Industry', 'Agriculture']].tolist()
    
    #add either consumption or production of district heating from fuel refinery 
    if FR_DH >= 0:
        DH.append(FR_DH)
        FR.append(0)
    else: 
        FR.append(abs(FR_DH))
        DH.append(0)
    
    
    #Flows out of the fuel refinery (node 11), electricity production (12) and district heating (13)
    source += [11]*10 + [12]*7 + [13]*5
    target += [12, 12, 13, 14, 15, 16, 17, 18, 12, 13,  # refinery: to electricity (incl. CHP), DH boilers, transport, sectors, and the electricity and heat it produces
               13, 14, 15, 16, 17, 18, 11,              # electricity: to district heating (heat pumps, electric boilers, CHP heat), transport, sectors, refinery
               15, 16, 17, 18, 11]                      # district heating: to the sectors and the refinery
    values += FR + El + DH
    color_link += ['rgba(33, 115, 120, 0.65)']*10 + ['rgba(175, 85, 0, 0.65)']*7 + ['rgba(220, 47, 2, 0.65)']*5
    color_node = ['#C5F567', '#A96CCE', '#5e60ce', '#FF7400', '#38ACAE', '#FEAA00', '#66E067', '#E467AC', '#FFFF32', '#4266BC', '#83D2F4', '#217378', '#AF5500', '#dc2f02'] + [c_dark]*5
        
    #data to dict, dict to sankey
    link = dict(source = source, target = target, value = values, color = color_link)
    node = dict(label = label, pad=20, thickness=15, color = color_node)
    data = go.Sankey(link = link, node = node, arrangement = "snap")
    
    return data

#The Sankey images are made with Kaleido, which needs Google Chrome and sometimes hangs. If making an image fails
#or takes longer than SANKEY_IMAGE_TIMEOUT seconds, a button that opens the diagram in the browser is shown instead.
SANKEY_IMAGE_TIMEOUT = 60
sankey_images_work = True #set to False after the first failure, so pySTREAM does not wait for Kaleido again

def save_sankey_image(fig, file_name):
    """Save the figure as a PNG image with Kaleido. Returns True if that worked within SANKEY_IMAGE_TIMEOUT."""
    global sankey_images_work
    if not sankey_images_work:
        return False
    image_fig = go.Figure(fig) #a copy, so a Kaleido that hangs does not hold on to the figure shown in the browser
    result = {}
    def save():
        try:
            image_fig.write_image(figure_path(file_name), scale=SCALE) #the image fills the graph area; its text is scaled like the rest
            result['saved'] = True
        except Exception as error:
            result['error'] = error
    #Save in a background thread, so pySTREAM can stop waiting if Kaleido hangs
    thread = threading.Thread(target=save, daemon=True)
    thread.start()
    thread.join(SANKEY_IMAGE_TIMEOUT)
    if not result.get('saved'):
        sankey_images_work = False
        print('The Sankey diagram could not be saved as an image:', result.get('error', 'Kaleido did not finish in time'))
    return result.get('saved', False)

def show_sankey(fig, file_name, frame):
    """Show a Sankey diagram in the frame as an image; clicking it opens the interactive diagram in the browser.
    If no image can be made, a button that opens the browser is shown instead."""
    def sankey_web():
        fig.update_layout(paper_bgcolor='white', width=1500, height=720)
        fig.show()
    if save_sankey_image(fig, file_name):
        sankey_image = ImageTk.PhotoImage(Image.open(figure_path(file_name)))
        sankey_button = FlatButton(frame, image=sankey_image, command=sankey_web, bd=0)
        sankey_button.image = sankey_image
        tip = "Tip: Click the image to open the graph in interactive mode in your browser!"
    else:
        sankey_button = FlatButton(frame, text="Open the energy flow (Sankey) diagram in your browser", command=sankey_web, bd=0, bg=c_lightish, fg=c_dark, activebackground=c_dark, activeforeground=c_white, font=font_button)
        tip = "The diagram could not be shown here (this needs Google Chrome), but it can be opened in your browser."
    sankey_button.place(x=ww*0.005, y=wh*0.04, width = graph_width, height = graph_height)
    sankey_tip = tk.Label(frame, text = tip, font=font_text, bd=0, bg=c_light, fg=c_dark)
    sankey_tip.place(x=ww*0.005, y=wh*0.58, width=graph_width, height=wh*0.02)

def sankey_SC():
    dataSC = sankey_diagram(df_fuelType_consumption_SC, df_Fc_SC, graphs_sc_frame)
    sankeyFig_SC = go.Figure(dataSC)
    sankeyFig_SC.update_layout(title_text="Energy streams - " + str(Scenario_Year), title_font=dict(size=20, family="Arial", color=c_dark), title_x=0.5, paper_bgcolor=c_light, width=graph_width/SCALE+1, height=graph_height/SCALE+1)
    show_sankey(sankeyFig_SC, "sankeySC.png", graphs_sc_frame)

def sankey_BY():
    dataBY = sankey_diagram(df_fuelType_consumption_BY, df_Fc_BY, graphs_sc_frame)
    sankeyFig_BY = go.Figure(dataBY)
    sankeyFig_BY.update_layout(title_text="Energy streams - " + str(Base_Year), title_font=dict(size=20, family="Arial", color=c_dark), title_x=0.5, paper_bgcolor=c_light, width=graph_width/SCALE+1, height=graph_height/SCALE+1)
    show_sankey(sankeyFig_BY, "sankeyBY.png", graphs_by_frame)

#Sankey button
def button_sankey_action():
    sankey_SC()
    sankey_BY()

button_sankey = FlatButton(graphs_frame, command=button_sankey_action, text='Sankey', bd=0, bg=c_lightish, fg=c_dark, font=font_button, activebackground=c_dark, activeforeground=c_white)
button_sankey.place(x=ww*0.1*0, y=wh*0.09, height=wh*0.07, width=ww*0.09)
#------------------------------------------------------------------------------

#Stacked bar plot of consumption in  sectors: Residential, Tertiary, Industry, Agriculture
#------------------------------------------------------------------------------
def FEC_stackedBar(df_FEC, year, frame):
    #Create stacked bar plot showing FEC for base year
    fig, ax = plot_stacked_bar(df_FEC, "Technologies", 4, (6, 3))
    ax.set_title("Final Energy Consumption per Sector - "+ year)
    ax.set_ylabel("[TJ]")
    canvas_FEC = FigureCanvasTkAgg(fig, master=frame)
    canvas_FEC.get_tk_widget().place(x=ww*0.005, y=wh*0.04, width = graph_width, height = graph_height)
    imageName = figure_path("FECstacked" + year + ".png")
    fig.savefig(imageName)

def button_FEC_stackedBar_action():
    FEC_stackedBar(df_FEC_BY, str(Base_Year), graphs_by_frame)
    FEC_stackedBar(df_FEC_SC, str(Scenario_Year), graphs_sc_frame)
    
button_FEC_stackedBar = FlatButton(graphs_frame, wraplength=ww*0.08, command=button_FEC_stackedBar_action, text='Sectors Technologies', bd=0, bg=c_lightish, fg=c_dark, font=font_button, activebackground=c_dark, activeforeground=c_white)
button_FEC_stackedBar.place(x=ww*0.1*1, y=wh*0.09, height=wh*0.07, width=ww*0.09)
#------------------------------------------------------------------------------

#Bar plot of consumption in all sectors: Residential, Tertiary, Industry, Agriculture, Transport
#------------------------------------------------------------------------------

def FEC_bar(df_FEC, df_transport_consumption, year, frame):
    sectors_consumption = df_FEC.iloc[:,:4].sum()
    sectors_consumption.loc['Transport'] = df_transport_consumption.sum().values[0]
    fig, ax = plot_bar(sectors_consumption*0.001)
    ax.set_title("Final Energy Consumption per Sector - "+ year)
    ax.tick_params(axis='x', labelrotation=0)
    canvas_FEC = FigureCanvasTkAgg(fig, master=frame)
    canvas_FEC.get_tk_widget().place(x=ww*0.005, y=wh*0.04, width = graph_width, height = graph_height)
    imageName = figure_path("FEC" + year + ".png")
    fig.savefig(imageName)

def button_FEC_bar_action():
    FEC_bar(df_FEC_BY, df_transport_consumption_BY, str(Base_Year), graphs_by_frame)
    FEC_bar(df_FEC_SC, df_transport_consumption_SC, str(Scenario_Year), graphs_sc_frame)
    
button_FEC_stackedBar = FlatButton(graphs_frame, wraplength=ww*0.08, command=button_FEC_bar_action, text='All sectors Consumption', bd=0, bg=c_lightish, fg=c_dark, font=font_button, activebackground=c_dark, activeforeground=c_white)
button_FEC_stackedBar.place(x=ww*0.1*2, y=wh*0.09, height=wh*0.07, width=ww*0.09)
#------------------------------------------------------------------------------

#Bar plot of CO2 in all sectors: Residential, Tertiary, Industry, Agriculture, Transport
#------------------------------------------------------------------------------

def CO2_bar(df_fuelType_CO2, year, frame):
    CO2_sums = df_fuelType_CO2.sum()
    # Rename specific rows
    new_names = {
                'Electricity production no CCS': 'Electricity no CCS',
                'Electricity production w. CCS': 'Electricity w. CCS',
                'Total emissions incl. CCS': 'Total w. CCS',
                'Total emissions excl. CCS': 'Total no CCS',
                }
    CO2_sums = CO2_sums.rename(new_names)
    fig, ax = plot_bar(CO2_sums)
    ax.set_ylabel("mill tons")
    ax.set_title("CO2 emissions per Sector - "+ year)
    canvas_FEC = FigureCanvasTkAgg(fig, master=frame)
    canvas_FEC.get_tk_widget().place(x=ww*0.005, y=wh*0.04, width = graph_width, height = graph_height)
    imageName = figure_path("CO2bar" + year + ".png")
    fig.savefig(imageName)

def button_CO2_bar_action():
    CO2_bar(df_fuelType_CO2_BY, str(Base_Year), graphs_by_frame)
    CO2_bar(df_fuelType_CO2_SC, str(Scenario_Year), graphs_sc_frame)
    
button_CO2_bar = FlatButton(graphs_frame, wraplength=ww*0.08, command=button_CO2_bar_action, text='CO2', bd=0, bg=c_lightish, fg=c_dark, font=font_button, activebackground=c_dark, activeforeground=c_white)
button_CO2_bar.place(x=ww*0.1*3, y=wh*0.09, height=wh*0.07, width=ww*0.09)
#------------------------------------------------------------------------------

#Bar plot of Electricity production for each fuel type
#------------------------------------------------------------------------------

def electricity_bar(df_Power_amount, year, frame):
    fig, ax = plot_bar(df_Power_amount)
    ax.tick_params(axis='x', labelrotation=90)
    ax.set_xlabel('')
    ax.set_title("Electricity production per production method - "+ year)
    canvas_FEC = FigureCanvasTkAgg(fig, master=frame)
    canvas_FEC.get_tk_widget().place(x=ww*0.005, y=wh*0.04, width = graph_width, height = graph_height)
    imageName = figure_path("electricityBar" + year + ".png")
    fig.savefig(imageName)

def button_electricity_bar_action():
    electricity_bar(df_Power_BY['Amount'], str(Base_Year), graphs_by_frame)
    electricity_bar(df_Power_SC['Amount'], str(Scenario_Year), graphs_sc_frame)
    
button_electricity_bar = FlatButton(graphs_frame, wraplength=ww*0.08, command=button_electricity_bar_action, text='Electricity', bd=0, bg=c_lightish, fg=c_dark, font=font_button, activebackground=c_dark, activeforeground=c_white)
button_electricity_bar.place(x=ww*0.1*4, y=wh*0.09, height=wh*0.07, width=ww*0.09)

#------------------------------------------------------------------------------

#Bar plot of District heat production for each fuel type
#------------------------------------------------------------------------------

def heat_bar(df_Heat_amount, year, frame):
    fig, ax = plot_bar(df_Heat_amount)
    ax.set_xlabel('')
    ax.set_title("District heating production per production method - "+ year)
    ax.tick_params(axis='x', labelrotation=55)
    canvas_FEC = FigureCanvasTkAgg(fig, master=frame)
    canvas_FEC.get_tk_widget().place(x=ww*0.005, y=wh*0.04, width = graph_width, height = graph_height)
    imageName = figure_path("heatBar" + year + ".png")
    fig.savefig(imageName)

def button_heat_bar_action():
    heat_bar(df_Heat_BY['Amount'], str(Base_Year), graphs_by_frame)
    heat_bar(df_Heat_SC['Amount'], str(Scenario_Year), graphs_sc_frame)
    
button_heat_bar = FlatButton(graphs_frame, wraplength=ww*0.08, command=button_heat_bar_action, text='District heat', bd=0, bg=c_lightish, fg=c_dark, font=font_button, activebackground=c_dark, activeforeground=c_white)
button_heat_bar.place(x=ww*0.1*5, y=wh*0.09, height=wh*0.07, width=ww*0.09)

#------------------------------------------------------------------------------

#Bar plot of resource consumption
#------------------------------------------------------------------------------
def primaryFuels_bar(df_fuelType_total, year, frame):
    df_fuelType_total = df_fuelType_total.drop(['Electricity', 'District Heating'])
    fig, ax = plot_bar(df_fuelType_total)
    ax.set_xlabel('')
    ax.set_title("Total primary fuel consumption - " + year)
    ax.tick_params(axis='x', labelrotation=55)
    canvas_FEC = FigureCanvasTkAgg(fig, master=frame)
    canvas_FEC.get_tk_widget().place(x=ww*0.005, y=wh*0.04, width = graph_width, height = graph_height)
    imageName = figure_path("primaryFuelBar" + year + ".png")
    fig.savefig(imageName)

def button_primaryFuels_bar_action():
    primaryFuels_bar(df_fuelType_consumption_BY['Total primary fuel consumption'], str(Base_Year), graphs_by_frame)
    primaryFuels_bar(df_fuelType_consumption_SC['Total primary fuel consumption'], str(Scenario_Year), graphs_sc_frame)
    
button_primaryFuels_bar = FlatButton(graphs_frame, wraplength=ww*0.08, command=button_primaryFuels_bar_action, text='Fuel consumption', bd=0, bg=c_lightish, fg=c_dark, font=font_button, activebackground=c_dark, activeforeground=c_white)
button_primaryFuels_bar.place(x=ww*0.1*6, y=wh*0.09, height=wh*0.07, width=ww*0.0920)

#------------------------------------------------------------------------------

#==============================================================================

#Create consumption frame that appears when consumption button is pressed
#==============================================================================
consumption_frame = tk.Frame(window)
def show_consumption_frame():
    consumption_frame.config(width=ww*0.8, height=wh*0.85, bg=c_white) 
    consumption_frame.place(x=ww*0.17, y=wh*0.025)
    heatConsumption_table()
    transportFuels_table()
    allocation_table()
    utilisation_table()
    
header_frame(consumption_frame, "Edit consumption", "Scenario", ww*0.18)

#Heating consumption frame
#containing editable heat consumption table for scenario year, button to save changes and sum row
#------------------------------------------------------------------------------
heat_frame = tk.Frame(consumption_frame, width=ww*0.43, height=wh*0.31, bg=c_light)
heat_frame.place(x=0, y=wh*0.076)

header_heat1 = tk.Label(heat_frame, text="Heating", font=font_subsubheader, bg=c_light, fg=c_dark)
header_heat1.place(x=ww*0.005, y=wh*0.008)
header_heat2 = tk.Label(heat_frame, text="energy consumption", font=font_subsubheader, bg=c_light, fg=c_mid)
header_heat2.place(x=ww*0.062, y=wh*0.008)

#Creating editable heating consumption table for scenario year
y_pos_consumption=wh*0.05
x_pos=ww*0.005

def heatConsumption_table():
    columns_heat = (list(df_beta_SC))
    columns_heat.insert(0,'Fuels')
    
    fuels_heat = ['Electricity', 'District heating', 'Coal', 'Oil', 'Natural gas', 'Biomass', 'Solar', 'Electricity']
    table_heat = TreeviewEdit(heat_frame, parent_frame=heat_frame, y_pos=y_pos_consumption, x_pos=x_pos, columns=columns_heat, unit='%')
    rows_heat = df_beta_SC.iloc[1:].values.astype(float)
    rows_heat = np.column_stack((fuels_heat, rows_heat))
    row_index_heat = list(df_beta_SC[1:].index.values)
    
    #Insert first column with technology types
    table_heat.heading("#0", text="    Technology Type", anchor='w')
    table_heat.column("#0", stretch=False, width=int(ww*0.081))
    column_width = [int(ww*0.08), int(ww*0.065), int(ww*0.065), int(ww*0.065), int(ww*0.065)]
    
    #insert columns and row content in table
    for column in columns_heat:
        table_heat.column(column, width=column_width[columns_heat.index(column)], stretch=False)
        table_heat.heading(column, text=column, anchor='w')
    for i in range(len(rows_heat)):
        values = rows_heat[i]
        formatted_values = ["{:.2%}".format(value) for value in values[1:].astype(float)]  # Format each element as a percentage
        table_heat.insert("", 'end', values=[values[0]] + formatted_values, text=row_index_heat[i])
    table_heat.place(x=x_pos, y=y_pos_consumption, width=ww*0.419, height=wh*0.192)

    #Creating button to save heating consumption table contents to dataframe
    def save_heat_action():
        df_saved_heat = save_table_to_dataframe(table_heat, has_children=False)
        
        #Remove the percentage sign and convert to decimal format before updating df_beta_SC
        #The saved table is indexed by fuel name while df_beta_SC is indexed by tech type,
        #so the rows are written by position (.iloc) instead of by label. A chained
        #assignment (df_beta_SC[sector][1:] = ...) writes to a temporary copy and is
        #silently discarded under pandas' copy-on-write.
        for sector in SECTORS:
            new_shares = df_saved_heat[sector].str.rstrip('%').astype(float) / 100
            df_beta_SC.iloc[1:, df_beta_SC.columns.get_loc(sector)] = new_shares.to_numpy()
        
        #Update sums
        heat_sums()
       
    save_heat_button = FlatButton(heat_frame, text="Save heat changes" , command=save_heat_action, bd=0, bg=c_accent, fg=c_dark, activebackground=c_dark, activeforeground=c_accent, font=font_text_bold)
    save_heat_button.place(x=ww*0.344, y=wh*0.245, height=wh*0.021, width=ww*0.08)

#Creating heat "table" with relevant sums
def heat_sums():
    label_heat_sum = tk.LabelFrame(heat_frame, text=" Sum", font=font_text_bold, labelanchor='sw', bd=0, bg=c_mid, fg=c_white)
    label_heat_sum.place(x=ww*0.005, y=wh*0.27, height=wh*0.021, width=ww*0.161)
    
    heat_sums = [df_beta_SC[SECTORS[0]][1:].sum(), df_beta_SC[SECTORS[1]][1:].sum(), df_beta_SC[SECTORS[2]][1:].sum(), df_beta_SC[SECTORS[3]][1:].sum()]
    label_heat_sum_value = [0, 0, 0, 0]
    
    for i in range(len(label_heat_sum_value)):
        label_heat_sum_value[i] = tk.LabelFrame(heat_frame, text="{:.2%}".format(heat_sums[i]), font=('Arial', pt(9),  'italic'), labelanchor='sw', bd=0, bg=c_white, fg=c_black)
        label_heat_sum_value[i].place(x=ww*0.166+int(ww*0.065)*i, y=wh*0.27, height=wh*0.021, width=int(ww*0.065))
   
    
    #Check if the sums add up to 100%
    errors_found = False
    for i in range(len(heat_sums)):
        if round(heat_sums[i], 2) != 1:
            label_heat_sum_value[i].config(fg='red') #make error sums that are more or less than 100% red
            errors_found = True
    if errors_found:
        tk.messagebox.showerror("Sum error","The sum for each sector in heating consumption should be 100%. Please correct this before updating scenario.")

heat_sums()

#------------------------------------------------------------------------------

#Transport fuels consumption frame
#containing editable transport fuel consumption table for scenario year, button to save changes and sum columns
#------------------------------------------------------------------------------
transportFuels_frame = tk.Frame(consumption_frame, width=ww*0.59, height=wh*0.4, bg=c_light)
transportFuels_frame.place(x=0, y=wh*0.4)

header_transportFuels1 = tk.Label(transportFuels_frame, text="Transport fuels", font=font_subsubheader, bg=c_light, fg=c_dark)
header_transportFuels1.place(x=ww*0.005, y=wh*0.008)
header_transportFuels2 = tk.Label(transportFuels_frame, text="energy consumption", font=font_subsubheader, bg=c_light, fg=c_mid)
header_transportFuels2.place(x=ww*0.115, y=wh*0.008)

#Creating editable transport fuels consumption table for scenario year
def transportFuels_table():
    columns_transportFuels = list(df_TransportFuel_SC)
    columns_transportFuels.insert(0,'')
    table_transportFuels = TreeviewEdit(transportFuels_frame, parent_frame=transportFuels_frame, y_pos=y_pos_consumption, x_pos=x_pos, columns=columns_transportFuels, show='headings', unit='%')
    show_group_rows(table_transportFuels)
    rows_transportFuels = df_TransportFuel_SC.values.astype(float)
    rows_transportFuels = np.delete(rows_transportFuels, 4, axis=0)
    row_index_transportFuels = list(df_TransportFuel_SC.index.values)
    row_index_transportFuels.remove('Bike') #Bike does not use fuels and is therefor removed
    rows_transportFuels = np.column_stack((row_index_transportFuels, rows_transportFuels))
    
    column_width = [int(ww*0.087), int(ww*0.05), int(ww*0.046), int(ww*0.042), int(ww*0.05), int(ww*0.048), int(ww*0.046), int(ww*0.048), int(ww*0.05), int(ww*0.056)] 
    
    #Creating empty parent rows to seperate table rows by transport sector type
    person_row = table_transportFuels.insert(parent='', index='end', values=('Person','','','','','','','','',''), open=True, tags=GROUP)
    goods_row =  table_transportFuels.insert(parent='', index='end', values=('Goods','','','','','','','','',''), open=True, tags=GROUP)
    fisheryAgriculture_row =  table_transportFuels.insert(parent='', index='end', values=('Fishery & Agriculture','','','','','','','','',''), open=True, tags=GROUP)
    
    #insert columns and row content in table
    for column in columns_transportFuels:
        table_transportFuels.column(column, width=column_width[columns_transportFuels.index(column)], stretch=False)
        table_transportFuels.heading(column, text=column, anchor='w')
    for i in range(len(rows_transportFuels)):
        values = rows_transportFuels[i]
        formatted_values = ["{:.2%}".format(value) for value in values[1:].astype(float)]  # Format each element as a percentage
        if i < 4:
            parent = person_row
        elif i < 8:
            parent = goods_row
        else:
            parent = fisheryAgriculture_row
        table_transportFuels.insert(parent=parent, index='end', values=[values[0]] + formatted_values)
    table_transportFuels.place(x=x_pos, y=y_pos_consumption, width=ww*0.52, height=wh*0.3)

    #Creating button to save transport fuel consumption table contents to dataframe
    def save_transportFuels_action():
        #in python tkinter I have a treeview table that has three parent items, each with a few children. I have a function that I use to turn my tables into dataframes, that I use for several tables. The other tables I use it for do not have any children, so it works well for them. However, for this table the output I get is only the parents, but I only want the children. How can I put the table into my function and only get the children into the dataframe
        df_saved_transportFuels = save_table_to_dataframe(table_transportFuels, has_children=True)
        
        #Remove the percentage sign and convert to decimal format before updating df_beta_SC
        for fuel in columns_transportFuels[1:]:
            df_TransportFuel_SC[fuel] = df_saved_transportFuels[fuel].str.rstrip('%').astype(float) / 100
        
        #Update sums
        transportFuels_sums()
       
    save_transportFuels_button = FlatButton(transportFuels_frame, text="Save transport fuel changes" , command=save_transportFuels_action, bd=0, bg=c_accent, fg=c_dark, activebackground=c_dark, activeforeground=c_accent, font=font_text_bold)
    save_transportFuels_button.place(x=ww*0.395, y=wh*0.355, height=wh*0.021, width=ww*0.13)


#Table next to the transport fuels table with the sum of the fuel shares of each transport type
table_transportFuels_sum = side_table(transportFuels_frame, 'Sum', int(ww*0.05), 'Sums.Treeview')
table_transportFuels_sum.place(x=ww*0.535, y=y_pos_consumption, height=wh*0.3)
show_group_rows(table_transportFuels_sum)
table_transportFuels_sum.tag_configure('error', foreground='red') #sums that are not 100%

def transportFuels_sums():
    table_transportFuels_sum.delete(*table_transportFuels_sum.get_children())
    errors_found = False
    for i, fuel_sum in enumerate(df_TransportFuel_SC.drop('Bike').sum(axis=1)): #Bike does not use fuels
        if i in (0, 4, 8): #next to the rows Person, Goods and Fishery & Agriculture
            table_transportFuels_sum.insert('', 'end', values=[''], tags=GROUP)
        error = round(fuel_sum, 2) != 1
        errors_found = errors_found or error
        table_transportFuels_sum.insert('', 'end', values=["{:.2%}".format(fuel_sum)], tags=('error',) if error else ())

    #Check if the sums add up to 100%
    if errors_found:
        tk.messagebox.showerror("Sum error","The sum of fuel for each transport type should be 100%. Please correct this before updating scenario.")

transportFuels_sums()
#------------------------------------------------------------------------------

#Transportation details frame
#containing editable Allocation of transport table (Person and Goods), editable Utilisation degree table (Person and Goods), and a save changes button for each table.
#------------------------------------------------------------------------------
transportDetails_frame = tk.Frame(consumption_frame, width=ww*0.1894, height=wh*0.725, bg=c_light)
transportDetails_frame.place(x=ww*0.6, y=wh*0.076)

header_transportDetails1 = tk.Label(transportDetails_frame, text="Transportation", font=font_subsubheader, bg=c_light, fg=c_dark)
header_transportDetails1.place(x=ww*0.005, y=wh*0.002)
header_transportDetails2 = tk.Label(transportDetails_frame, text="details", font=font_subsubheader, bg=c_light, fg=c_mid)
header_transportDetails2.place(x=ww*0.112, y=wh*0.002)

#Creating editable Allocation of transport table (Person and Goods) 
def allocation_table():
    columns_transportAllocation = [list(df_TransportWork_SC)[0]]
    columns_transportAllocation.insert(1,'[pkm] or [tkm]')
    table_transportAllocation = TreeviewEdit(transportDetails_frame, parent_frame=transportDetails_frame, y_pos=y_pos_consumption, x_pos=x_pos, columns=columns_transportAllocation, show='headings', unit='%')
    show_group_rows(table_transportAllocation)
    
    column_width = [int(ww*0.11), int(ww*0.068)]
    #Person rows
    rows_transportAllocationPerson = df_TransportWork_SC['Allocation of transport work'][:5]
    row_index_transportAllocationPerson = list(df_TransportWork_SC.index.values)[:5]
    rows_transportAllocationPerson = np.column_stack((row_index_transportAllocationPerson, rows_transportAllocationPerson))
    
    #insert columns content in table
    for i in columns_transportAllocation:
        table_transportAllocation.column(i, width=column_width[columns_transportAllocation.index(i)], stretch=False)
        table_transportAllocation.heading(i, text=columns_transportAllocation[columns_transportAllocation.index(i)], anchor='w')
    #insert row content in table
    person_parent = table_transportAllocation.insert("", 'end', values=["Person", ""], open=True, tags=GROUP)
    for i in range(len(rows_transportAllocationPerson)):
        values = rows_transportAllocationPerson[i]
        formatted_values = [values[0]] + ["{:.2%}".format(value) for value in values[1:].astype(float)]  # Format each element as a percentage
        table_transportAllocation.insert(person_parent, 'end', values=formatted_values)
        
    #Goods rows
    rows_transportAllocationGoods = df_TransportWork_SC['Allocation of transport work'][5:]
    row_index_transportAllocationGoods = list(df_TransportWork_SC.index.values)[5:]
    rows_transportAllocationGoods = np.column_stack((row_index_transportAllocationGoods, rows_transportAllocationGoods))
    
    #insert row content in table
    goods_parent = table_transportAllocation.insert("", 'end', values=["Goods", ""], open=True, tags=GROUP)
    for i in range(len(rows_transportAllocationGoods)):
        values = rows_transportAllocationGoods[i]
        formatted_values = [values[0]] + ["{:.2%}".format(value) for value in values[1:].astype(float)]  # Format each element as a percentage
        table_transportAllocation.insert(goods_parent, 'end', values=formatted_values)
    
    table_transportAllocation.place(x=x_pos, y=y_pos_consumption, height=wh*0.255)

    #Creating button to save Allocation of transport work table contents to dataframe
    def save_transportAllocation_action():
        df_saved_transportAllocation= save_table_to_dataframe(table_transportAllocation, has_children=True)
        #Remove the percentage sign and convert to decimal format before updating df_beta_SC
        df_TransportWork_SC['Allocation of transport work'] = df_saved_transportAllocation['[pkm] or [tkm]'].str.rstrip('%').astype(float) / 100
        
        #Update sums
        transportAllocation_sums()
       
    save_transportAllocation_button = FlatButton(transportDetails_frame, text="Save transport work changes" , command=save_transportAllocation_action, bd=0, bg=c_accent, fg=c_dark, activebackground=c_dark, activeforeground=c_accent, font=font_text_bold)
    save_transportAllocation_button.place(x=ww*0.052, y=y_pos_consumption+wh*0.26, height=wh*0.02, width=ww*0.13)


#Create a "table" of relevant sums
def transportAllocation_sums():
    label_AllocationPerson_sum = tk.LabelFrame(transportDetails_frame, text="Person sum", font=font_text_bold, bd=0, bg=c_mid, fg=c_white, labelanchor='s')
    label_AllocationPerson_sum. place(x=x_pos, y=y_pos_consumption+wh*0.285, height=wh*0.02, width=ww*0.052)
    allocation_sums=[0,0]
    label_Allocation_sum_value = ['']*2
    allocation_sums[0] = df_TransportWork_SC['Allocation of transport work'][:5].sum()
    label_Allocation_sum_value[0] = tk.LabelFrame(transportDetails_frame, text="{:.2%}".format(allocation_sums[0]), font=('Arial', pt(9), 'italic'), bd=0, bg=c_white, fg=c_black, labelanchor='s')
    label_Allocation_sum_value[0].place(x=x_pos+ww*0.052, y=y_pos_consumption+wh*0.285, height=wh*0.02, width=ww*0.038)
    
    label_AllocationGoods_sum = tk.LabelFrame(transportDetails_frame, text="Goods sum", font=font_text_bold, bd=0, bg=c_mid, fg=c_white, labelanchor='s')
    label_AllocationGoods_sum. place(x=x_pos+ww*0.09, y=y_pos_consumption+wh*0.285, height=wh*0.02, width=ww*0.052)
    
    allocation_sums[1] = df_TransportWork_SC['Allocation of transport work'][5:].sum()
    label_Allocation_sum_value[1] = tk.LabelFrame(transportDetails_frame, text="{:.2%}".format(allocation_sums[1]), font=('Arial', pt(9), 'italic'), bd=0, bg=c_white, fg=c_black, labelanchor='s')
    label_Allocation_sum_value[1].place(x=x_pos+ww*0.14, y=y_pos_consumption+wh*0.285, height=wh*0.02, width=ww*0.038)
    
    pad_frame1 = tk.Frame(transportDetails_frame, bg=c_mid, bd=0)
    pad_frame1.place(x=x_pos, y=y_pos_consumption+wh*0.305, width=ww*0.14, height=px(3))
    pad_frame2 = tk.Frame(transportDetails_frame, bg=c_white, bd=0)
    pad_frame2.place(x=x_pos+ww*0.052, y=y_pos_consumption+wh*0.305, width=ww*0.038, height=px(3))
    pad_frame3 = tk.Frame(transportDetails_frame, bg=c_white, bd=0)
    pad_frame3.place(x=x_pos+ww*0.14, y=y_pos_consumption+wh*0.305, width=ww*0.038, height=px(3))
    
    #Check if the sums add up to 100%
    errors_found = False
    for i in range(len(allocation_sums)):
        if round(allocation_sums[i], 2) != 1:
            label_Allocation_sum_value[i].config(fg='red') #make error sums that are more or less than 100% red
            errors_found = True
    if errors_found:
        tk.messagebox.showerror("Sum error","The sum for person and goods in Allocation of transport work should each be 100%. Please correct this before updating scenario.")
    
transportAllocation_sums()

#Creating editable Utilisation degree table (Person and Goods) 
def utilisation_table():
    columns_transportUtilisation = [list(df_TransportWork_SC)[1]]
    columns_transportUtilisation.insert(1,'[pkm] or [tkm]')
    table_transportUtilisation = TreeviewEdit(transportDetails_frame, parent_frame=transportDetails_frame, y_pos=y_pos_consumption+315, x_pos=x_pos, columns=columns_transportUtilisation, show='headings', unit='%', allowed=(1, 100), what='The utilisation degree (how full the vehicles are on average)')
    show_group_rows(table_transportUtilisation)
    column_width = [int(ww*0.11), int(ww*0.068)]
    
    #Person rows
    rows_transportUtilisationPerson = df_TransportWork_SC['Utilisation degree'][:5]
    row_index_transportUtilisationPerson = list(df_TransportWork_SC.index.values)[:5]
    rows_transportUtilisationPerson = np.column_stack((row_index_transportUtilisationPerson, rows_transportUtilisationPerson))
    
    #insert columns content in table
    for i in columns_transportUtilisation:
        table_transportUtilisation.column(i, width=column_width[columns_transportUtilisation.index(i)], stretch=False)
        table_transportUtilisation.heading(i, text=columns_transportUtilisation[columns_transportUtilisation.index(i)], anchor='w')
    #insert row content in table
    person_parent = table_transportUtilisation.insert("", 'end', values=["Person", ""], open=True, tags=GROUP)
    for i in range(len(rows_transportUtilisationPerson)):
        values = rows_transportUtilisationPerson[i]
        formatted_values = [values[0]] + ["{:.2%}".format(value) for value in values[1:].astype(float)]  # Format each element as a percentage
        table_transportUtilisation.insert(person_parent, 'end', values=formatted_values)
        
    #Goods rows
    rows_transportUtilisationGoods = df_TransportWork_SC['Utilisation degree'][5:]
    row_index_transportUtilisationGoods = list(df_TransportWork_SC.index.values)[5:]
    rows_transportUtilisationGoods = np.column_stack((row_index_transportUtilisationGoods, rows_transportUtilisationGoods))
    
    #insert row content in table
    goods_parent = table_transportUtilisation.insert("", 'end', values=["Goods", ""], open=True, tags=GROUP)
    for i in range(len(rows_transportUtilisationGoods)):
        values = rows_transportUtilisationGoods[i]
        formatted_values = [values[0]] + ["{:.2%}".format(value) for value in values[1:].astype(float)]  # Format each element as a percentage
        table_transportUtilisation.insert(goods_parent, 'end', values=formatted_values)
    
    table_transportUtilisation.place(x=x_pos, y=y_pos_consumption+wh*0.35, height=wh*0.255)

    #Creating button to save Utilisation degree table contents to dataframe
    def save_transportUtilisation_action():
        df_saved_transportUtilisation= save_table_to_dataframe(table_transportUtilisation, has_children=True)
        #Remove the percentage sign and convert to decimal format before updating df_beta_SC
        df_TransportWork_SC['Utilisation degree'] = df_saved_transportUtilisation['[pkm] or [tkm]'].str.rstrip('%').astype(float) / 100
           
    save_transportUtilisation_button = FlatButton(transportDetails_frame, text="Save utilisation degree changes" , command=save_transportUtilisation_action, bd=0, bg=c_accent, fg=c_dark, activebackground=c_dark, activeforeground=c_accent, font=font_text_bold)
    save_transportUtilisation_button.place(x=ww*0.048, y=y_pos_consumption+wh*0.61, height=wh*0.02, width=ww*0.135)


#------------------------------------------------------------------------------

#==============================================================================

#Create production frame that appears when production button is pressed
#==============================================================================

production_frame = tk.Frame(window)
def show_production_frame():
    production_frame.config(width=ww*0.8, height=wh*0.85, bg=c_white) 
    production_frame.place(x=ww*0.17, y=wh*0.025)
    electricity_table()
    heat_table()
    resources_table()
    plots_production()
    
    
header_frame(production_frame, "Edit production", "Scenario", ww*0.16)

heading_production = ["Production method", "Shares"]

#Electricity production edit frame
#containing production table for scenario year, button to save changes and table with sums
#------------------------------------------------------------------------------
electricity_frame = tk.Frame(production_frame, width=ww*0.2045, height=wh*0.725, bg=c_light)
electricity_frame.place(x=0, y=wh*0.076)

header_electricity2 = tk.Label(electricity_frame, text='Production', font=font_subsubsubheader, bg=c_light, fg=c_mid)
header_electricity2.place(x=ww*0.005, y=wh*0.035, anchor='nw')
header_electricity1 = tk.Label(electricity_frame, text='Electricity', font=font_subsubheader, bg=c_light, fg=c_dark)
header_electricity1.place(x=ww*0.005, y=wh*0.005, anchor='nw')

#Creating editable Electricity production table for scenario year
columns_electricity = ['Tech Types', 'Shares']
y_pos_production = wh*0.076
x_pos = ww*0.005

def electricity_table():
    table_electricity = TreeviewEdit(electricity_frame, parent_frame=electricity_frame, y_pos=y_pos_production, x_pos=x_pos, columns=columns_electricity, show='headings', unit='%')
    rows_electricity = df_Power_SC[['Tech Types', 'Shares']][:-1].to_numpy().tolist()
    widths_p = [int(ww*0.11), int(ww*0.04)]
    #Insert columns and row content in table
    for i in columns_electricity:
        table_electricity.column(i, width=widths_p[columns_electricity.index(i)], stretch=False)
        table_electricity.heading(i, text=heading_production[columns_electricity.index(i)], anchor='w')
    for data in rows_electricity:
        v = data
        formatted_v = [v[0]] + ["{:.2%}".format(value) for value in v[1:]]  # Format each element as a percentage
        table_electricity.insert("", 'end', iid=v[0], values=formatted_v)
    table_electricity.place(x=x_pos, y=y_pos_production, height=wh*0.36)
    
    #Creating button to save electricity table contents to dataframe
    def save_electricity_action():
        df_saved_electricity = save_table_to_dataframe(table_electricity, has_children=False)
    
        #Remove the percentage sign and convert to decimal format before updating df_Power_SC
        df_Power_SC['Shares'] = df_saved_electricity['Shares'].str.rstrip('%').astype(float) / 100

        #Set electricity imports (the last row is not in the table): they cover what the other shares do not
        #Rounded so that floating-point rest values (like 1e-16) are not treated as real imports
        df_Power_SC.loc[df_Power_SC.index[-1], 'Shares'] = round(1 - df_Power_SC.iloc[:-1]['Shares'].sum(), 10)

        df_Power_SC['Amount'] = df_Power_SC['Shares'] * total_electricity_consumption_SC
        #Update sums
        electricity_sums()
        
        #Update amounts produced
        electricity_amounts()
        
    save_electricity_action()    
        
   
    save_electricity_button = FlatButton(electricity_frame, text="Save electricity changes" , command=save_electricity_action, bd=0, bg=c_accent, fg=c_dark, activebackground=c_dark, activeforeground=c_accent, font=font_text_bold)
    save_electricity_button.place(x=ww*0.089, y=wh*0.44, height=wh*0.025, width=ww*0.11)

#Creating electricity "table" with relevant sums
def electricity_sums():
    #demand coverage
    label_demand_coverage = tk.LabelFrame(electricity_frame, text=" Demand coverage", font=font_text_bold, labelanchor='sw', bd=0, bg=c_mid, fg=c_white)
    label_demand_coverage.place(x=ww*0.005, y=wh*0.47, height=wh*0.023, width=int(ww*0.105))
    demand_coverage = df_Power_SC['Shares'][:-1].sum()
    label_demand_coverage_value = tk.LabelFrame(electricity_frame, text="{:.2%}".format(demand_coverage), font=font_text, labelanchor='s', bd=0, bg=c_white, fg=c_black)
    label_demand_coverage_value.place(x=int(ww*0.11), y=wh*0.47, height=wh*0.023, width=int(ww*0.09))
    
    #electricity imports (shares below 100 %) or exports (shares above 100 %)
    electricity_imports = round(1 - demand_coverage, 10)
    df_Power_SC.loc['Imported electricity', 'Shares'] = electricity_imports
    label_electricity_imports = tk.LabelFrame(electricity_frame, text=" Electricity imports" if electricity_imports >= 0 else " Electricity exports", font=font_text_bold, labelanchor='sw', bd=0, bg=c_mid, fg=c_white)
    label_electricity_imports.place(x=ww*0.005, y=wh*0.47+wh*0.023*1, height=wh*0.023, width=int(ww*0.105))
    label_electricity_imports_value = tk.LabelFrame(electricity_frame, text="{:.2%}".format(abs(electricity_imports)), font=font_text, labelanchor='s', bd=0, bg=c_white, fg=c_black)
    label_electricity_imports_value.place(x=int(ww*0.11), y=wh*0.47+wh*0.023*1, height=wh*0.023, width=int(ww*0.09))
    
    #total demand
    label_total_demand = tk.LabelFrame(electricity_frame, text=" Total demand", font=font_text_bold, labelanchor='sw', bd=0, bg=c_mid, fg=c_white)
    label_total_demand.place(x=ww*0.005, y=wh*0.47+wh*0.023*2, height=wh*0.023, width=int(ww*0.105))
    total_demand = round(total_electricity_consumption_SC,2)
    label_total_demand_value = tk.LabelFrame(electricity_frame, text=(str(total_demand) + ' PJ'), font=font_text, labelanchor='s', bd=0, bg=c_white, fg=c_black)
    label_total_demand_value.place(x=int(ww*0.11), y=wh*0.47+wh*0.023*2, height=wh*0.023, width=int(ww*0.09))
    
    #total production
    label_total_production = tk.LabelFrame(electricity_frame, text=" Total production", font=font_text_bold, labelanchor='sw', bd=0, bg=c_mid, fg=c_white)
    label_total_production.place(x=ww*0.005, y=wh*0.47+wh*0.023*3, height=wh*0.023, width=int(ww*0.105))
    total_production = round(df_Power_SC['Amount'][:-1].sum(),2)
    label_total_production_value = tk.LabelFrame(electricity_frame, text=(str(total_production) + ' PJ'), font=font_text, labelanchor='s', bd=0, bg=c_white, fg=c_black)
    label_total_production_value.place(x=int(ww*0.11), y=wh*0.47+wh*0.023*3, height=wh*0.023, width=int(ww*0.09))
    
    #Pad frames are empty frames added to make it look better
    pad_frame1 = tk.Frame(electricity_frame, bg=c_mid, bd=0)
    pad_frame1.place(x=ww*0.005, y=wh*0.47+wh*0.023*4, width=int(ww*0.105), height=px(5))
    pad_frame2 = tk.Frame(electricity_frame, bg=c_white, bd=0)
    pad_frame2.place(x=int(ww*0.11), y=wh*0.47+wh*0.023*4, width=int(ww*0.09), height=px(5))
    
#electricity_sums()

def show_amounts(table, amounts):
    table.delete(*table.get_children())
    for amount in amounts:
        table.insert('', 'end', values=[str(round(amount, 2))])

#Table next to the electricity shares table with the amounts of electricity produced per production unit type in [PJ]
table_amountsEl = side_table(electricity_frame, 'Amount [PJ]', int(ww*0.045), 'Amounts.Treeview')
table_amountsEl.place(x=ww*0.154, y=y_pos_production, height=wh*0.36)
def electricity_amounts():
    show_amounts(table_amountsEl, df_Power_SC['Amount'][:-1]) #the last row (imports) is not in the shares table
        
#electricity_amounts()   
    
#------------------------------------------------------------------------------

#District heating production edit frame
#containing production table for scenario year, button to save changes and table with sums
#------------------------------------------------------------------------------
DH_frame = tk.Frame(production_frame, width=ww*0.18, height=wh*0.725, bg=c_light)
DH_frame.place(x=ww*0.215, y=wh*0.076)

header_DH1 = tk.Label(DH_frame, text='Production', font=font_subsubsubheader, bg=c_light, fg=c_mid)
header_DH1.place(x=ww*0.005, y=wh*0.035, anchor='nw')
header_DH2 = tk.Label(DH_frame, text='District heat boiler', font=font_subsubheader, bg=c_light, fg=c_dark)
header_DH2.place(x=ww*0.005, y=wh*0.005, anchor='nw')

#Creating editable DH production table for scenario year
columns_heat = ['Tech Types', 'Shares']

def heat_table():
    table_DH= TreeviewEdit(DH_frame, parent_frame=DH_frame, y_pos=y_pos_production, x_pos=x_pos, columns=columns_heat, show='headings', unit='%')
    rows_heat = df_Heat_SC[['Tech Types', 'Shares']][:-1].to_numpy().tolist()
    widths_h = [int(ww*0.09), int(ww*0.035)]
    #Insert columns and row content in table
    for i in columns_heat:
        table_DH.column(i, width=widths_h[columns_heat.index(i)], stretch=False)
        table_DH.heading(i, text=heading_production[columns_heat.index(i)], anchor='w')
    for data in rows_heat:
        v = data 
        formatted_v = [v[0]] + ["{:.2%}".format(value) for value in v[1:]]  # Format each element as a percentage
        table_DH.insert("", 'end', iid=v[0], values=formatted_v) 
    table_DH.place(x=x_pos, y=y_pos_production, height=wh*0.232)
    

    #Creating button to save DH table contents to dataframe
    def save_DH_action():
        df_saved_DH = save_table_to_dataframe(table_DH, has_children=False)
        
        #Remove the percentage sign and convert to decimal format before updating df_Heat_SC
        df_Heat_SC['Shares'] = df_saved_DH['Shares'].str.rstrip('%').astype(float) / 100
        
        #Set heat import
        df_Heat_SC.loc[df_Heat_SC.index[-1], 'Shares'] = round(1-df_Heat_SC.iloc[:-1]['Shares'].sum(), 10)        
        df_Heat_SC.loc['Heat imports', 'Amount'] = total_DH_demand_SC * df_Heat_SC.loc['Heat imports', 'Shares']
        
        #calculate demand for DH boiler production. If the CHP production covers the total DH demand, there is no DH boiler production
        if df_Power_SC['CHP heat production'].sum() < total_DH_demand_SC: 
            total_DH_boiler = total_DH_demand_SC - df_Power_SC['CHP heat production'].sum()
        else:
            total_DH_boiler = 0
            
        #Set DH boiler production amount
        df_Heat_SC['Amount'] = df_Heat_SC['Shares'] * total_DH_boiler
        
        #Update sums
        DH_sums()
        #Update amounts produced
        DH_amounts()
    
    save_DH_action()
    
    save_DH_button = FlatButton(DH_frame, text="Save district heat changes" , command=save_DH_action, bd=0, bg=c_accent, fg=c_dark, activebackground=c_dark, activeforeground=c_accent, font=font_text_bold)
    save_DH_button.place(x=ww*0.055, y=wh*0.312, height=wh*0.025, width=ww*0.12)

#Creating "table" of relevant sums
def DH_sums():
    #heat imports (shares below 100 %) or exports (shares above 100 %)
    heat_imports_percent = df_Heat_SC.iloc[-1]['Shares']
    label_heat_imports = tk.LabelFrame(DH_frame, text=" Heat imports" if heat_imports_percent >= 0 else " Heat exports", font=font_text_bold, labelanchor='sw', bd=0, bg=c_mid, fg=c_white)
    label_heat_imports.place(x=ww*0.005, y=wh*0.344, height=wh*0.023, width=ww*0.09)
    label_heat_imports_percent = tk.LabelFrame(DH_frame, text="{:.2%}".format(abs(heat_imports_percent)), font=font_text, labelanchor='s', bd=0, bg=c_white, fg=c_black)
    label_heat_imports_percent.place(x=ww*0.095, y=wh*0.344, height=wh*0.023, width=ww*0.04)
    heat_imports = round(abs(df_Heat_SC.loc['Heat imports', 'Amount']),1)
    label_heat_imports_value = tk.LabelFrame(DH_frame, text=str(heat_imports) + " PJ", font=font_text, labelanchor='s', bd=0, bg=c_white, fg=c_black)
    label_heat_imports_value.place(x=ww*0.135, y=wh*0.344, height=wh*0.023, width=ww*0.04)
    
    #total production
    label_total_production = tk.LabelFrame(DH_frame, text=" Total production", font=font_text_bold, labelanchor='sw', bd=0, bg=c_mid, fg=c_white)
    label_total_production.place(x=ww*0.005, y=wh*0.344+wh*0.023, height=wh*0.023, width=ww*0.09)
    total_production_percent = df_Heat_SC['Shares'][:-1].sum()
    label_total_production_percent = tk.LabelFrame(DH_frame,  text="{:.2%}".format(total_production_percent), font=font_text, labelanchor='s', bd=0, bg=c_white, fg=c_black)
    label_total_production_percent.place(x=ww*0.095, y=wh*0.344+wh*0.023, height=wh*0.023, width=ww*0.04)
    total_production = round(df_Heat_SC['Amount'][:-1].sum(),1)
    label_total_production_value = tk.LabelFrame(DH_frame, text=str(total_production) + " PJ", font=font_text, labelanchor='s', bd=0, bg=c_white, fg=c_black)
    label_total_production_value.place(x=ww*0.135, y=wh*0.344+wh*0.023, height=wh*0.023, width=ww*0.04)
    
    #Pad frames are empty frames added to make it look better
    pad_frame1 = tk.Frame(DH_frame, bg=c_mid)
    pad_frame1.place(x=ww*0.005, y=wh*0.344+wh*0.023*2, height=px(5), width=ww*0.09)
    pad_frame2 = tk.Frame(DH_frame, bg=c_white)
    pad_frame2.place(x=ww*0.095, y=wh*0.344+wh*0.023*2, height=px(5), width=ww*0.08)
    
    #combined heat and power (CHP)
    label_chp = tk.LabelFrame(DH_frame, text=" Combined heat & power (CHP)", font=font_text_bold, labelanchor='sw', bd=0, bg=c_mid, fg=c_white)
    label_chp.place(x=ww*0.005, y=wh*0.344+wh*0.023*3, height=wh*0.023, width=ww*0.13)
    chp = round(df_Power_SC['CHP heat production'].sum(),1)
    labe_chp_value = tk.LabelFrame(DH_frame, text=str(chp) + " PJ", font=font_text, labelanchor='s', bd=0, bg=c_white, fg=c_black)
    labe_chp_value.place(x=ww*0.135, y=wh*0.344+wh*0.023*3, height=wh*0.023, width=ww*0.04)
    
    #Total district heat
    label_total_DH = tk.LabelFrame(DH_frame, text=" Total district heating", font=font_text_bold, labelanchor='sw', bd=0, bg=c_mid, fg=c_white)
    label_total_DH.place(x=ww*0.005, y=wh*0.344+wh*0.023*4, height=wh*0.023, width=ww*0.13)
    total_DH =  round(chp + total_production,1)
    labe_total_DH_value = tk.LabelFrame(DH_frame, text=str(total_DH) + " PJ", font=font_text, labelanchor='s', bd=0, bg=c_white, fg=c_black)
    labe_total_DH_value.place(x=ww*0.135, y=wh*0.344+wh*0.023*4, height=wh*0.023, width=ww*0.04)
    
    #Pad frames are empty frames added to make it look better
    pad_frame3 = tk.Frame(DH_frame, bg=c_mid)
    pad_frame3.place(x=ww*0.005, y=wh*0.344+wh*0.023*5, height=px(5), width=ww*0.13)
    pad_frame4 = tk.Frame(DH_frame, bg=c_white)
    pad_frame4.place(x=ww*0.135, y=wh*0.344+wh*0.023*5, height=px(5), width=ww*0.04)
    
    
    
#DH_sums()

#Table next to the district heat boiler shares table with the amounts of heat produced per production unit type in [PJ]
table_amountsDH = side_table(DH_frame, 'Amount [PJ]', int(ww*0.045), 'Amounts.Treeview')
table_amountsDH.place(x=ww*0.129, y=y_pos_production, height=wh*0.232)
def DH_amounts():
    show_amounts(table_amountsDH, df_Heat_SC['Amount'][:-1]) #the last row (heat imports) is not in the shares table

#DH_amounts()
        
def DH_piePlot():
    pie_data = [abs(df_fuelType_consumption_SC.loc['District Heating']['Heat production']), abs(df_fuelType_consumption_SC.loc['District Heating']['DH boilers'])]
    labels = ['CHP heat', 'DH boilers']
    if df_fuelType_consumption_SC.loc['District Heating']['Fuel Refinery'] < 0:
        labels.append('Fuel refinery')
        pie_data.append(abs(df_fuelType_consumption_SC.loc['District Heating']['Fuel Refinery']))
    pie_data = pd.Series(pie_data, index=labels)
    fig, ax = plot_pie(pie_data, [PLOT_COLORS[6],PLOT_COLORS[4],PLOT_COLORS[3]])
    canvas_FEC = FigureCanvasTkAgg(fig, master=DH_frame)
    canvas_FEC.get_tk_widget().place(x=ww*0.029 , y=wh*0.491, width = ww*0.125, height = ww*0.125) 
    label_DH_piePlot = tk.LabelFrame(DH_frame, text='District heat production sources', font=('Arial', pt(9), 'bold'), bd=0, bg=c_light, fg=c_mid)
    label_DH_piePlot.place(x=ww*0.025, y=wh*0.473, height=wh*0.0213, width = ww*0.18)
#------------------------------------------------------------------------------

#Resources available frame
#containing available resources table, and graphs showing available vs used for scenario year and base year.
#------------------------------------------------------------------------------
resources_frame = tk.Frame(production_frame, width=ww*0.5, height=wh*0.41, bg=c_light)
resources_frame.place(x=ww*0.405, y=wh*0.076)


header_resources1 = tk.Label(resources_frame, text='Available & used', font=font_subsubsubheader, bg=c_light, fg=c_mid)
header_resources1.place(x=ww*0.005, y=wh*0.035, anchor='nw')
header_resources2 = tk.Label(resources_frame, text='Resources', font=font_subsubheader, bg=c_light, fg=c_dark)
header_resources2.place(x=ww*0.005, y=wh*0.005, anchor='nw')

df_Resources_SC

def fix_resources_rows():
    rows_resources = df_Resources_SC[['Type', 'Amount']].to_numpy()
    rows_resources = rows_resources[rows_resources[:, 0].argsort()]
    rows_resources2 = df_Fc_p_SC['Total primary fuel consumption'].drop(['Electricity', 'District Heating', 'Imported electricity', 'Imported heat'])
    oil_combined = rows_resources2['Oil'] + rows_resources2['Crude Oil']
    rows_resources2['Oil'] = oil_combined
    rows_resources2 = rows_resources2.drop('Crude Oil')
    rows_resources2 = rows_resources2.sort_index()

    rows_resources = np.column_stack((rows_resources, rows_resources2))
    return rows_resources


def resources_table():
    rows_resources = fix_resources_rows()
    table_resources_heading = [' Resource', ' Available', ' Used']
    columns_resources = ['Type', 'Amount', 'Total primary fuel consumption']
    table_resources = ttk.Treeview(resources_frame, columns=columns_resources, show='headings')
    widths = [int(ww*0.1), int(ww*0.04), int(ww*0.03)]
    #Insert columns and row content in table
    for i in range(len(table_resources_heading)):
        table_resources.column(columns_resources[i], width=widths[i], stretch=False)
        table_resources.heading(columns_resources[i], text=table_resources_heading[i], anchor='w')
    for data in rows_resources:
        v = data 
        formatted_v = [v[0]] + ["{:.0f}".format(value) for value in v[1:]]  # Format each element as a percentage
        table_resources.insert("", 'end', iid=v[0], values=formatted_v) 
    table_resources.place(x=ww*0.005, y=wh*0.076, height=wh*0.317)

def resources_plot():
    rows_resources = fix_resources_rows()
    # Extract names, available, and used
    names = rows_resources[:, 0]
    available_data = rows_resources[:, 1].astype(float)
    used_data = rows_resources[:, 2].astype(float)
    indices = np.arange(len(names))  # Create an array of indices for plotting
    bar_width = 0.4
    fig, ax = plt.subplots()
    plot_standards(fig, ax)
    
    # Create bars for available and used
    bar_available = ax.barh(indices, available_data[::-1], bar_width, label='Available', color=c_mid)
    bar_used = ax.barh(indices + bar_width, used_data[::-1], bar_width, label='Used', color=c_dark)

    # Bars longer than the axis (500 PJ, see below) are cut off; write their value at the end of the bar
    for bars, data in [(bar_available, available_data[::-1]), (bar_used, used_data[::-1])]:
        for bar, value in zip(bars, data):
            if value > 500:
                ax.text(495, bar.get_y() + bar.get_height() / 2, '{:.0f} PJ >'.format(value), ha='right', va='center', fontsize=7, color=c_white, weight='bold')

    # Set the names as labels for the bars
    ax.set_yticks(indices + bar_width)
    ax.set_yticklabels(names[::-1])  # reverse the order of the names
    ax.set_xticks([])
    ax.set_xlim(0, 500)  # limit the length of the bars to make small bars easier to see
    
    # Add labels and a legend
    ax.set_title('Resources - available and used [PJ]', color=c_dark, weight='bold')
    
    fig.legend(frameon=False, loc='upper center', ncol=2, bbox_to_anchor=(0.5, 0.9))
    fig.patch.set_facecolor(c_light)
    ax.spines['bottom'].set_visible(False)
    # Set the background color of the axes to be transparent
    ax.set_facecolor(c_light)
    
    canvas_resources = FigureCanvasTkAgg(fig, master=resources_frame)
    canvas_resources.get_tk_widget().place(x=ww*0.143, y=wh*0.032, width=ww*0.26, height=wh*0.423)
    
resources_plot()
#------------------------------------------------------------------------------

#Create production plots
#------------------------------------------------------------------------------
#Create pie plot of fuel categories in electricity production
productionPlots_frame = tk.Frame(production_frame, width=ww*0.5, height=wh*0.3, bg=c_light)
productionPlots_frame.place(x=ww*0.405, y=wh*0.501)

def plots_production():
    electricity_c_piePlot()    
    DH_c_piePlot()
    DH_piePlot()
        
def electricity_c_piePlot():
    pie_data = df_fuelCategory_consumption_SC.iloc[1:]['Electricity production']
    fig, ax = plot_pie(pie_data, PLOT_COLORS[0:3])
    canvas_FEC = FigureCanvasTkAgg(fig, master=productionPlots_frame)
    canvas_FEC.get_tk_widget().place(x=ww*0.05, y=wh*0.05, width = ww*0.127, height = ww*0.127)
    label_electricity_c_piePlot = tk.LabelFrame(productionPlots_frame, text='Electricity production categories', font=('Arial', pt(9), 'bold'), bd=0, bg=c_light, fg=c_dark)
    label_electricity_c_piePlot.place(x=ww*0.05, y=wh*0.02, height=wh*0.0213, width = ww*0.18)
    
def DH_c_piePlot():
    pie_data = df_fuelCategory_consumption_SC.iloc[1:3]['Heat production'] + df_fuelCategory_consumption_SC.iloc[1:3]['DH boilers']
    fig, ax = plot_pie(pie_data, PLOT_COLORS[0:3])
    ax.set_title('District heat production categories incl. CHP', fontsize=9, color=c_black, weight='bold')
    canvas_FEC = FigureCanvasTkAgg(fig, master=productionPlots_frame)
    canvas_FEC.get_tk_widget().place(x=ww*0.22, y=wh*0.05, width = ww*0.127, height = ww*0.127)    
    label_electricity_c_piePlot = tk.LabelFrame(productionPlots_frame, text='District heat production categories incl. CHP', font=('Arial', pt(9), 'bold'), bd=0, bg=c_light, fg=c_dark)
    label_electricity_c_piePlot.place(x=ww*0.2, y=wh*0.02, height=wh*0.0213, width = ww*0.18)

#------------------------------------------------------------------------------

#==============================================================================

#Functions to create similar tables and plots for production of electricity and heat for scenario and base year 
#==============================================================================
def create_production_tables(frame, df_Power, df_Heat, column, form):
    heading_production = ["Production method", "Produced"]
    #Creating Power production table
    columns_powerProduction = ['Tech Types', column]
    table_powerProduction = ttk.Treeview(frame, columns=columns_powerProduction, show='headings')
    rows_powerProduction = df_Power[['Tech Types', column]].to_numpy().tolist()
    power_shares = df_Power[column]
    widths_p = [int(ww*0.105), int(ww*0.04)]
    #Insert columns and row content in table
    for i in columns_powerProduction:
        table_powerProduction.column(i, width=widths_p[columns_powerProduction.index(i)], stretch=False)
        table_powerProduction.heading(i, text=heading_production[columns_powerProduction.index(i)], anchor='w')
    for data in rows_powerProduction:
        v = data
        formatted_v = [v[0]] + [form.format(value) for value in v[1:]]  # Format each element 
        table_powerProduction.insert("", 'end', iid=v[0], values=formatted_v)
    table_powerProduction.place(x=ww*0.015, y=wh*0.068, height=wh*0.382)
    
    #Creating Heat production table
    columns_heatProduction = ['Tech Types', column]
    table_heatProduction = ttk.Treeview(frame, columns=columns_heatProduction, show='headings')
    rows_heatProduction = df_Heat[['Tech Types', column]].to_numpy().tolist()
    widths_h = [int(ww*0.075), int(ww*0.04)]
    #Insert columns and row content in table
    for i in columns_heatProduction:
        table_heatProduction.column(i, width=widths_h[columns_heatProduction.index(i)], stretch=False)
        table_heatProduction.heading(i, text=heading_production[columns_heatProduction.index(i)], anchor='w')
    for data in rows_heatProduction:
        v = data 
        formatted_v = [v[0]] + [form.format(value) for value in v[1:]]  # Format each element
        table_heatProduction.insert("", 'end', iid=v[0], values=formatted_v) 
    table_heatProduction.place(x=ww*0.2, y=wh*0.068, height=wh*0.253)


def electricity_hbar(df_Power_amount, year, frame):
    fig, ax = plot_horizontal_bar(df_Power_amount)
    ax.set_ylabel('')
    ax.tick_params(axis='y', labelsize=7)
    ax.set_title("Electricity production per production method           ", fontsize=9)
    canvas_FEC = FigureCanvasTkAgg(fig, master=frame)
    canvas_FEC.get_tk_widget().place(x=ww*0.001, y=wh*0.47, width = ww*0.22, height = wh*0.25)
   
def heat_hbar(df_Heat_amount, year, frame):
    fig, ax = plot_horizontal_bar(df_Heat_amount)
    ax.set_ylabel('')
    ax.tick_params(axis='y', labelsize=7)
    ax.set_title("District heating production per production method            ", fontsize=9)
    canvas_FEC = FigureCanvasTkAgg(fig, master=frame)
    canvas_FEC.get_tk_widget().place(x=ww*0.22, y=wh*0.47, width = ww*0.18 , height = wh*0.25)
    
#==============================================================================

#Create "Base Year Data" frame that appears when Base year button is pressed
#==============================================================================
by_frame = tk.Frame(window)
def show_by_frame():
    by_frame.config(width=ww*0.8, height=wh*0.85, bg=c_white) 
    by_frame.place(x=ww*0.17, y=wh*0.025)
    create_FEC_BY_frame()
    create_production_BY_frame()
    
header_frame(by_frame, "Base Year Data", str(Base_Year), ww*0.16)

#Final Energy Consumption (FEC) base year frame
#containing FEC graphs and table
#------------------------------------------------------------------------------

FEC_BY_frame = tk.Frame(by_frame, width=ww*0.385, height=wh*0.725, bg=c_light)
FEC_BY_frame.place(x=0, y=wh*0.076)

header_FEC = tk.Label(FEC_BY_frame, text="Final Energy Consumption (FEC)", font=font_subsubheader, bg=c_light, fg=c_dark)
header_FEC.place(x=ww*0.005, y=wh*0.001, width=ww*0.385)

def create_FEC_BY_frame():

    #Creating Final Energy Consumption (FEC) in [TJ] table for base year 
    def create_FEC_table():
        table_FEC = ttk.Treeview(FEC_BY_frame, columns=list(df_FEC_BY))
        rows_FEC = df_FEC_BY.to_numpy()
        rows_FEC = list(np.around(np.array(rows_FEC),1))
        row_index = list(df_FEC_BY.index.values)
        
        table_FEC.heading("#0", text="    Technology Type", anchor='w')
        table_FEC.column("#0", stretch=False, width=int(ww*0.1))
        #insert columns and row content in table
        for i in list(df_FEC_BY):
            table_FEC.column(i, width=int(ww*0.055), stretch=False)
            table_FEC.heading(i, text=i, anchor='w')
        for j in range(len(rows_FEC)):
            v = [r for r in rows_FEC[j]] 
            table_FEC.insert("", 'end', values=v, text=row_index[j])
        table_FEC.place(x=ww*0.005, y=wh*0.5, height=wh*0.21)
        
    create_FEC_table()
    
    #button to switch table to FEC [TJ]
    FEC_table_button = FlatButton(FEC_BY_frame, text="FEC [TJ]", command=create_FEC_table, bd=0, bg=c_midlight, fg=c_white, activebackground=c_dark, activeforeground=c_white, font=font_text_bold)
    FEC_table_button.config(width=7, height=1)
    FEC_table_button.place(x=ww*0.005, y=wh*0.47)
    
    
    #Creating Final Energy Consumption (FEC) in percent table for base year
    def create_beta_table():
        table_beta = ttk.Treeview(FEC_BY_frame, columns=list(df_beta_BY))
        rows_beta = df_beta_BY.to_numpy()
        rows_beta = list(np.array(rows_beta))
        row_index = list(df_beta_BY.index.values)
        
        table_beta.heading("#0", text="    Technology Type", anchor='w')
        table_beta.column("#0", stretch=False, width=int(ww*0.1))
        #insert columns and row content in table
        for i in list(df_beta_BY):
            table_beta.column(i, width=int(ww*0.055), stretch=False)
            table_beta.heading(i, text=i, anchor='w')
        for j in range(len(rows_beta)):
            v = rows_beta[j]
            formatted_v = ["{:.2%}".format(value) for value in v]  # Format each element as a percentage
            table_beta.insert("", 'end', values=formatted_v, text=row_index[j])
        table_beta.place(x=ww*0.005, y=wh*0.5, height=wh*0.21)
        
    #button to switch table to % shares
    shares_table_button = FlatButton(FEC_BY_frame, text="%", command=create_beta_table, bd=0, bg=c_midlight, fg=c_white, activebackground=c_dark, activeforeground=c_white, font=font_text_bold)
    shares_table_button.config(width=3, height=1)
    shares_table_button.place(x=ww*0.047, y=wh*0.47)
    

    #Create stacked bar plot showing FEC for base year
    fig, ax = plot_stacked_bar(df_FEC_BY, "Technologies", 4, (6, 3))
    ax.set_title("Final Energy Consumption per Sector - "+ str(Base_Year))
    ax.set_ylabel("[TJ]")
    canvas_FEC = FigureCanvasTkAgg(fig, master=FEC_BY_frame)
    canvas_FEC.get_tk_widget().place(x=ww*0.005, y=wh*0.039, width=ww*0.375, height=wh*0.43)
#------------------------------------------------------------------------------

#Energy production base year frame
#containing energy production, power and heat table
#------------------------------------------------------------------------------
productionBY_frame = tk.Frame(by_frame, width=ww*0.395, height=wh*0.725, bg=c_light)
productionBY_frame.place(x=ww*0.395, y=wh*0.076)

header_production = tk.Label(productionBY_frame, text="Production", font=font_subsubheader, bg=c_light, fg=c_dark)
header_production.place(y=wh*0.001, x=0, width=ww*0.36)
header_production_power = tk.Label(productionBY_frame, text="Power", font=font_subsubsubheader, bg=c_light, fg=c_mid)
header_production_power.place(y=wh*0.035, x=ww*0.015, anchor='nw')
header_production_heat = tk.Label(productionBY_frame, text="District Heating", font=font_subsubsubheader, bg=c_light, fg=c_mid)
header_production_heat.place(y=wh*0.035, x=ww*0.2, anchor='nw')

def production_hbar_plot_BY():
    electricity_hbar(df_Power_BY['Amount'], str(Base_Year), productionBY_frame)
    heat_hbar(df_Heat_BY['Amount'], str(Base_Year), productionBY_frame)

def create_production_BY_frame():
    create_production_tables(productionBY_frame, df_Power_BY, df_Heat_BY, 'Amount', '{:.2f}')
    
    #Horizontal bar plot of power and heat production
    production_hbar_plot_BY()
    
#button to switch production table to PJ
production_table_buttonBY = FlatButton(productionBY_frame, text="PJ", command=create_production_BY_frame, bd=0, bg=c_midlight, fg=c_white, activebackground=c_dark, activeforeground=c_white, font=font_text_bold)
production_table_buttonBY.config(width=2, height=1)
production_table_buttonBY.place(x=ww*0.225, y=wh*0.006)

    
def create_productionShares_BY_frame():
    #Creating Power production table for Base year
    df_Power_BY['Shares'] = [amount/df_Power_BY['Amount'].sum() for amount in df_Power_BY['Amount']]
    df_Heat_BY['Shares'] = [amount/df_Heat_BY['Amount'].sum() for amount in df_Heat_BY['Amount']]

    create_production_tables(productionBY_frame, df_Power_BY, df_Heat_BY, 'Shares', "{:.2%}")
    
#button to switch production table to % shares
productionShares_table_buttonBY = FlatButton(productionBY_frame, text="%", command=create_productionShares_BY_frame, bd=0, bg=c_midlight, fg=c_white, activebackground=c_dark, activeforeground=c_white, font=font_text_bold)
productionShares_table_buttonBY.config(width=2, height=1)
productionShares_table_buttonBY.place(x=ww*0.24, y=wh*0.006)


#------------------------------------------------------------------------------


#==============================================================================

#Create Scenario data frame that appears when Scenario button is pressed
#==============================================================================

scenario_frame = tk.Frame(window)
def show_scenario_frame():
    scenario_frame.config(width=ww*0.8, height=wh*0.85, bg=c_white) 
    scenario_frame.place(x=ww*0.17, y=wh*0.025)
    create_FEC_SC_frame()
    create_production_SC_frame()
    header_frame(scenario_frame, "Scenario Year Data", str(Scenario_Year), ww*0.2)
    
#header_frame(scenario_frame, "Scenario Year Data", str(Scenario_Year), ww*0.2)
    
#Final Energy Consumption (FEC) scenario year frame, containing FEC graphs and table
#------------------------------------------------------------------------------
FEC_SC_frame = tk.Frame(scenario_frame, width=ww*0.385, height=wh*0.725, bg=c_light)
FEC_SC_frame.place(x=0, y=wh*0.076)

header_FEC_SC = tk.Label(FEC_SC_frame, text="Final Energy Consumption (FEC)", font=font_subsubheader, bg=c_light, fg=c_dark)
header_FEC_SC.place(x=ww*0.005, y=wh*0.001, width=ww*0.385)

def create_FEC_SC_frame():
#Creating Final Energy Consumption (FEC) table for scenario year

    #table containing FEC values in TJ
    def create_FEC_table():
        table_FEC = ttk.Treeview(FEC_SC_frame, columns=list(df_FEC_SC))
        rows_FEC = df_FEC_SC.to_numpy()
        rows_FEC = list(np.around(np.array(rows_FEC),1))
        row_index = list(df_FEC_SC.index.values)
        
        table_FEC.heading("#0", text="    Technology Type", anchor='w')
        table_FEC.column("#0", stretch=False, width=int(ww*0.1))
        
        #Insert columns and row content in table
        for i in list(df_FEC_SC):
            table_FEC.column(i, width=int(ww*0.055), stretch=False)
            table_FEC.heading(i, text=i, anchor='w')
        for j in range(len(rows_FEC)):
            v = [r for r in rows_FEC[j]] 
            table_FEC.insert("", 'end', values=v, text=row_index[j])
        table_FEC.place(x=ww*0.005, y=wh*0.5, height=wh*0.21)
    
    create_FEC_table()
    
    #button to switch table to FEC [TJ]
    FEC_table_button = FlatButton(FEC_SC_frame, text="FEC [TJ]", command=create_FEC_table, bd=0, bg=c_midlight, fg=c_white, activebackground=c_dark, activeforeground=c_white, font=font_text_bold)
    FEC_table_button.config(width=7, height=1)
    FEC_table_button.place(x=ww*0.005, y=wh*0.47)

    #table containing FEC values as percentages % (beta = shares of consumption)
    def create_beta_table():
        table_beta = ttk.Treeview(FEC_SC_frame, columns=list(df_beta_SC))
        rows_beta = df_beta_SC.to_numpy()
        rows_beta = list(np.array(rows_beta))
        row_index = list(df_beta_SC.index.values)

        
        table_beta.heading("#0", text="    Technology Type", anchor='w')
        table_beta.column("#0", stretch=False, width=int(ww*0.1))
        
        #Insert columns and row content in table
        for i in list(df_beta_SC):
            table_beta.column(i, width=int(ww*0.055), stretch=False)
            table_beta.heading(i, text=i, anchor='w')
        for j in range(len(rows_beta)):
            v = rows_beta[j]
            formatted_v = ["{:.2%}".format(value) for value in v]  # Format each element as a percentage
            table_beta.insert("", 'end', values=formatted_v, text=row_index[j])
        table_beta.place(x=ww*0.005, y=wh*0.5, height=wh*0.21)
        
    #button to switch table to % shares
    shares_table_button = FlatButton(FEC_SC_frame, text="%", command=create_beta_table, bd=0, bg=c_midlight, fg=c_white, activebackground=c_dark, activeforeground=c_white, font=font_text_bold)
    shares_table_button.config(width=3, height=1)  
    shares_table_button.place(x=ww*0.047, y=wh*0.47)
          

    #Creating tkinter canvas to display FEC plot as figure        
    fig, ax = plot_stacked_bar(df_FEC_SC, "Technologies", 4, (6, 3))
    ax.set_title("Final Energy Consumption per Sector - "+ str(Scenario_Year))
    ax.set_ylabel("[TJ]")
    canvas_FEC_SC = FigureCanvasTkAgg(fig, master=FEC_SC_frame)
    canvas_FEC_SC.get_tk_widget().place(x=ww*0.005, y=wh*0.039, width=ww*0.375, height=wh*0.43)

#------------------------------------------------------------------------------

#Energy production scenario year frame
#containing energy production, power and heat table
#------------------------------------------------------------------------------
productionSC_frame = tk.Frame(scenario_frame, width=ww*0.395, height=wh*0.725, bg=c_light)
productionSC_frame.place(x=ww*0.395, y=wh*0.076)

header_production = tk.Label(productionSC_frame, text="Production", font=font_subsubheader, bg=c_light, fg=c_dark)
header_production.place(y=wh*0.001, x=0, width=ww*0.36)
header_production_power = tk.Label(productionSC_frame, text="Power", font=font_subsubsubheader, bg=c_light, fg=c_mid)
header_production_power.place(y=wh*0.035, x=ww*0.015, anchor='nw')
header_production_heat = tk.Label(productionSC_frame, text="District Heating", font=font_subsubsubheader, bg=c_light, fg=c_mid)
header_production_heat.place(y=wh*0.035, x=ww*0.2, anchor='nw')

def production_hbar_plot_SC():
    electricity_hbar(df_Power_SC['Amount'], str(Base_Year), productionSC_frame)
    heat_hbar(df_Heat_SC['Amount'], str(Base_Year), productionSC_frame)
    
def create_production_SC_frame():
    create_production_tables(productionSC_frame, df_Power_SC, df_Heat_SC, 'Amount', "{:.2f}")
    
    #Horizontal bar plot of power and heat production
    production_hbar_plot_SC()
    
#button to switch production table to PJ
production_table_buttonSC = FlatButton(productionSC_frame, text="PJ", command=create_production_SC_frame, bd=0, bg=c_midlight, fg=c_white, activebackground=c_dark, activeforeground=c_white, font=font_text_bold)
production_table_buttonSC.config(width=2, height=1)
production_table_buttonSC.place(x=ww*0.225, y=wh*0.006)

def create_productionShares_SC_frame():
    #Creating Power production table for scenario year
    create_production_tables(productionSC_frame, df_Power_SC, df_Heat_SC, 'Shares', "{:.2%}")
    
#button to switch production table to % shares
productionShares_table_buttonSC = FlatButton(productionSC_frame, text="%", command=create_productionShares_SC_frame, bd=0, bg=c_midlight, fg=c_white, activebackground=c_dark, activeforeground=c_white, font=font_text_bold)
productionShares_table_buttonSC.config(width=2, height=1)
productionShares_table_buttonSC.place(x=ww*0.24, y=wh*0.006)

#------------------------------------------------------------------------------

#==============================================================================

#Create Transport data frame that appears when Transport button is pressed
#==============================================================================

transport_frame = tk.Frame(window)
def show_transport_frame():
    transport_frame.config(width=ww*0.8, height=wh*0.85, bg=c_white) 
    transport_frame.place(x=ww*0.17, y=wh*0.025)
    transport_table_amounts_BY()
    transport_table_amounts_SC()
    total_fuel_transport_table_BY()
    total_fuel_transport_table_SC()
    transport_fuel_plot_BY()
    transport_fuel_plot_SC()
    
header_frame(transport_frame, "Transport data", "Base year & Scenario", ww*0.16)

#Functions used to create transport fuel consumption tables, amounts and shares for both years
#----------------------------------------------------------------------------------
def transport_table_shares(df_TransportFuel, frame):
    columns_transportFuels = list(df_TransportFuel)
    columns_transportFuels.insert(0,'')
    table_transportFuels_shares = ttk.Treeview(frame, columns=columns_transportFuels, show='headings')
    show_group_rows(table_transportFuels_shares)
    rows_transportFuels = df_TransportFuel.values.astype(float)
    rows_transportFuels = np.delete(rows_transportFuels, 4, axis=0)
    row_index_transportFuels = list(df_TransportFuel.index.values)
    row_index_transportFuels.remove('Bike') #Bike does not use fuels and is therefor removed
    rows_transportFuels = np.column_stack((row_index_transportFuels, rows_transportFuels))
    
    column_width = [int(ww*0.09), int(ww*0.042), int(ww*0.04), int(ww*0.031), int(ww*0.048), int(ww*0.038), int(ww*0.035), int(ww*0.039), int(ww*0.042), int(ww*0.053)] 
    
    #Creating empty parent rows to seperate table rows by transport sector type
    person_row = table_transportFuels_shares.insert(parent='', index='end', values=('Person','','','','','','','','',''), open=True, tags=GROUP)
    goods_row =  table_transportFuels_shares.insert(parent='', index='end', values=('Goods','','','','','','','','',''), open=True, tags=GROUP)
    fisheryAgriculture_row =  table_transportFuels_shares.insert(parent='', index='end', values=('Fishery & Agriculture','','','','','','','','',''), open=True, tags=GROUP)
    
    #insert columns and row content in table
    for column in columns_transportFuels:
        table_transportFuels_shares.column(column, width=column_width[columns_transportFuels.index(column)], stretch=False)
        table_transportFuels_shares.heading(column, text=column, anchor='w')
    for i in range(len(rows_transportFuels)):
        values = rows_transportFuels[i]
        formatted_values = ["{:.0%}".format(value) for value in values[1:].astype(float)]  # Format each element as a percentage
        if i < 4:
            parent = person_row
        elif i < 8:
            parent = goods_row
        else:
            parent = fisheryAgriculture_row
        table_transportFuels_shares.insert(parent=parent, index='end', values=[values[0]] + formatted_values)
    return table_transportFuels_shares

def transport_table_amounts(df_TransportFuel, frame):
    columns_transportFuels = list(df_TransportFuel)
    columns_transportFuels.insert(0,'')
    table_transportFuels_amounts = ttk.Treeview(frame, columns=columns_transportFuels, show='headings')
    show_group_rows(table_transportFuels_amounts)
    rows_transportFuels = df_TransportFuel.values.astype(float)
    rows_transportFuels = np.delete(rows_transportFuels, 4, axis=0)
    row_index_transportFuels = list(df_TransportFuel.index.values)
    row_index_transportFuels.remove('Bike') #Bike does not use fuels and is therefor removed
    rows_transportFuels = np.column_stack((row_index_transportFuels, rows_transportFuels))
    
    column_width = [int(ww*0.09), int(ww*0.042), int(ww*0.04), int(ww*0.031), int(ww*0.048), int(ww*0.038), int(ww*0.035), int(ww*0.039), int(ww*0.042), int(ww*0.053)] 
    
    #Creating empty parent rows to seperate table rows by transport sector type
    person_row = table_transportFuels_amounts.insert(parent='', index='end', values=('Person','','','','','','','','',''), open=True, tags=GROUP)
    goods_row =  table_transportFuels_amounts.insert(parent='', index='end', values=('Goods','','','','','','','','',''), open=True, tags=GROUP)
    fisheryAgriculture_row =  table_transportFuels_amounts.insert(parent='', index='end', values=('Fishery & Agriculture','','','','','','','','',''), open=True, tags=GROUP)
    
    #insert columns and row content in table
    for column in columns_transportFuels:
        table_transportFuels_amounts.column(column, width=column_width[columns_transportFuels.index(column)], stretch=False)
        table_transportFuels_amounts.heading(column, text=column, anchor='w')
    for i in range(len(rows_transportFuels)):
        values = rows_transportFuels[i]
        formatted_values = ["{:.0f}".format(value) for value in values[1:].astype(float)] 
        if i < 4:
            parent = person_row
        elif i < 8:
            parent = goods_row
        else:
            parent = fisheryAgriculture_row
        table_transportFuels_amounts.insert(parent=parent, index='end', values=[values[0]] + formatted_values)
    return table_transportFuels_amounts
#----------------------------------------------------------------------------------

#Total Fuel consumption table
#----------------------------------------------------------------------------------
def total_fuel_transport_table(df_transport_consumption_fuels, frame, r):
    heading_fuel= ["Used [TJ]", 'Share [%]']
    columns_fuel = ['Total transport consumption', 'Share of total transport consumption']
    
    table_total_fuel_transport = ttk.Treeview(frame, columns=columns_fuel)
    rows_total_fuel_transport = df_transport_consumption_fuels[columns_fuel][:r].to_numpy().tolist()
    row_index_fuel = df_transport_consumption_fuels[:r].index.tolist()
    
    table_total_fuel_transport.heading("#0", text="     Fuel", anchor='w')
    table_total_fuel_transport.column("#0", stretch=False, width=int(ww*0.06))
    
    #insert columns and row content in table
    for i in range(len(columns_fuel)):
        table_total_fuel_transport.column(columns_fuel[i], width=int(ww*0.04), stretch=False)
        table_total_fuel_transport.heading(columns_fuel[i], text=heading_fuel[i], anchor='w')
    for j in range(len(rows_total_fuel_transport)):
        v = rows_total_fuel_transport[j]
        formatted_v = ["{:.2%}".format(value) for value in v[1:]]  # Format each element as a percentage
        table_total_fuel_transport.insert("", 'end', values=[int(v[0])]+formatted_v, text=row_index_fuel[j])
    return table_total_fuel_transport

#----------------------------------------------------------------------------------

#Base year transport frame
#Containing fuel sum table and fuel per transport type table
#----------------------------------------------------------------------------------
transport_BY_frame = tk.Frame(transport_frame, width=ww*0.67, height=wh*0.355, bg=c_light)
transport_BY_frame.place(x=0, y=wh*0.076)

header_transport_BY = tk.Label(transport_BY_frame, text="Base year", font=font_subsubheader, bg=c_light, fg=c_dark)
header_transport_BY.place(x=ww*0.005, y=wh*0.001, anchor='nw')
header_transport_BYsub = tk.Label(transport_BY_frame, text="Fuel consumption", font=font_subsubsubheader, bg=c_light, fg=c_mid)
header_transport_BYsub.place(x=ww*0.08, y=wh*0.005, anchor='nw')

width_BY=ww*0.287

def transport_table_amounts_BY():
    df_TransportFuel_amounts_BY = df_TransportFuel_BY.mul(df_transport_consumption_BY['Energy consumption'], axis=0)
    table_transportFuels_amounts_BY = transport_table_amounts(df_TransportFuel_amounts_BY, transport_BY_frame)
    table_transportFuels_amounts_BY.place(x=ww*0.005, y=wh*0.04, width=width_BY, height=wh*0.303)
    
def transport_table_shares_BY():
    table_transportFuels_shares_BY = transport_table_shares(df_TransportFuel_BY, transport_BY_frame)
    table_transportFuels_shares_BY.place(x=ww*0.005, y=wh*0.04, width=width_BY, height=wh*0.303)

#button to switch table to amounts [TJ]
transport_amounts_table_button_BY = FlatButton(transport_BY_frame, text="TJ", command=transport_table_amounts_BY, bd=0, bg=c_midlight, fg=c_white, activebackground=c_dark, activeforeground=c_white, font=font_text_bold)
transport_amounts_table_button_BY.config(width=3, height=1)
transport_amounts_table_button_BY.place(x=ww*0.254, y=wh*0.007)


#button to switch table to percent
transport_shares_table_button_BY = FlatButton(transport_BY_frame, text="%", command=transport_table_shares_BY, bd=0, bg=c_midlight, fg=c_white, activebackground=c_dark, activeforeground=c_white, font=font_text_bold)
transport_shares_table_button_BY.config(width=3, height=1)
transport_shares_table_button_BY.place(x=ww*0.274, y=wh*0.007)

def total_fuel_transport_table_BY():
    table_total_fuel_transport_BY = total_fuel_transport_table(df_transport_consumption_fuels_BY, transport_BY_frame, 5)
    table_total_fuel_transport_BY.place(x=ww*0.3, y=wh*0.04, height=wh*0.13)

def transport_fuel_plot_BY():
    plot_data = df_TransportFuel_BY[df_TransportFuel_BY.index != 'Bike']
    xlabels = plot_data.index.tolist()
    plot_data = plot_data.transpose()
    fig, ax = plot_stacked_bar(plot_data, "Fuels", 10, (4.8, 2))
    ax.set_title("Transportation fuel shares - " + str(Base_Year))
    xlabels[xlabels.index('Trucks and cargo vans')] = 'Trucks'
    ax.set_xticklabels(xlabels)
    yvalues = [0.2, 0.4, 0.6, 0.8, 1]
    ax.set_yticks([0.2, 0.4, 0.6, 0.8, 1])
    ax.set_yticklabels(['{:,.0%}'.format(y) for y in yvalues])
    ax.set_xlabel("")
    ax.legend(fontsize=7, loc='center left', bbox_to_anchor=(1, 0.5), alignment='left', facecolor=c_light, framealpha=0)
    ax.tick_params(axis='x', labelsize=7, labelrotation=70)
    canvas_FEC_SC = FigureCanvasTkAgg(fig, master=transport_BY_frame)
    canvas_FEC_SC.get_tk_widget().place(x=ww*0.456, y=wh*0.01, width=ww*0.23, height=wh*0.344)
#----------------------------------------------------------------------------------

#Scenario year transport frame 
#Containing fuel sum table and fuel per transport type table
#----------------------------------------------------------------------------------
transport_SC_frame = tk.Frame(transport_frame, width=ww*0.867, height=wh*0.355, bg=c_light)
transport_SC_frame.place(x=0, y=wh*0.445)

header_transport_SC = tk.Label(transport_SC_frame, text="Scenario", font=font_subsubheader, bg=c_light, fg=c_dark)
header_transport_SC.place(x=ww*0.005, y=wh*0.001, anchor='nw')
header_transport_SCsub = tk.Label(transport_SC_frame, text="Fuel consumption", font=font_subsubsubheader, bg=c_light, fg=c_mid)
header_transport_SCsub.place(x=ww*0.082, y=wh*0.005, anchor='nw')

width_SC=ww*0.454   

def transport_table_amounts_SC():
    df_TransportFuel_amounts_SC = df_TransportFuel_SC.mul(df_transport_consumption_SC['Energy consumption'], axis=0)
    table_transportFuels_amounts_SC = transport_table_amounts(df_TransportFuel_amounts_SC, transport_SC_frame)
    table_transportFuels_amounts_SC.place(x=ww*0.005, y=wh*0.04, width=width_SC, height=wh*0.303)

def transport_table_shares_SC():
    table_transportFuels_SC = transport_table_shares(df_TransportFuel_SC, transport_SC_frame)
    table_transportFuels_SC.place(x=ww*0.005, y=wh*0.04, width=width_SC, height=wh*0.303)

#button to switch table to amounts [TJ]
transport_amounts_table_button_SC = FlatButton(transport_SC_frame, text="TJ", command=transport_table_amounts_SC, bd=0, bg=c_midlight, fg=c_white, activebackground=c_dark, activeforeground=c_white, font=font_text_bold)
transport_amounts_table_button_SC.config(width=3, height=1)
transport_amounts_table_button_SC.place(x=ww*0.42, y=wh*0.007)

#button to switch table to percent
transport_shares_table_button_SC = FlatButton(transport_SC_frame, text="%", command=transport_table_shares_SC, bd=0, bg=c_midlight, fg=c_white, activebackground=c_dark, activeforeground=c_white, font=font_text_bold)
transport_shares_table_button_SC.config(width=3, height=1)
transport_shares_table_button_SC.place(x=ww*0.44, y=wh*0.007)


def total_fuel_transport_table_SC():
    table_total_fuel_transport_SC = total_fuel_transport_table(df_transport_consumption_fuels_SC, transport_SC_frame, 9)
    table_total_fuel_transport_SC.place(x=ww*0.467, y=wh*0.04, height=wh*0.215)
    
    
def transport_fuel_plot_SC():
    plot_data = df_TransportFuel_SC[df_TransportFuel_SC.index != 'Bike']
    xlabels = plot_data.index.tolist()
    plot_data = plot_data.transpose()
    fig, ax = plot_stacked_bar(plot_data, "Fuels", 10, (4.8, 2))
    ax.set_title("Transportation fuel shares - " + str(Scenario_Year))
    xlabels[xlabels.index('Trucks and cargo vans')] = 'Trucks'
    ax.set_xticklabels(xlabels)
    yvalues = [0.2, 0.4, 0.6, 0.8, 1]
    ax.set_yticks([0.2, 0.4, 0.6, 0.8, 1])
    ax.set_yticklabels(['{:,.0%}'.format(y) for y in yvalues])
    ax.set_xlabel("")
    ax.legend(fontsize=7, loc='center left', bbox_to_anchor=(1, 0.5), alignment='left', facecolor=c_light, framealpha=0)
    ax.tick_params(axis='x', labelsize=7, labelrotation=70)
    canvas_FEC_SC = FigureCanvasTkAgg(fig, master=transport_SC_frame)
    canvas_FEC_SC.get_tk_widget().place(x=ww*0.61, y=-wh*0.07, width=ww*0.22, height=wh*0.47)
#----------------------------------------------------------------------------------


#==============================================================================

#Create CO2 data frame that appears when CO2 button is pressed
#==============================================================================

CO2_frame = tk.Frame(window)
def show_CO2_frame():
    CO2_frame.config(width=ww*0.8, height=wh*0.85, bg=c_white) 
    CO2_frame.place(x=ww*0.17, y=wh*0.025)
    CO2_table_BY()
    CO2_table_SC()

header_frame(CO2_frame, "CO2 data", "Base year & Scenario", px(190))

#Table with CO2 emissions per sector and fuel
#----------------------------------------------------------------------------------
def CO2_table(frame, df_fuel_CO2):
    heading_CO2 = ['Electricity no CCS', 'Electricity w. CCS', 'District Heat', 'Fuel Refinery', 'Residential', 'Tertiary', 'Industry', 'Agriculture', 'Transport']
    columns_CO2 = df_fuel_CO2.columns.tolist()
    columns_CO2 = columns_CO2[0:-2]
    table_CO2 = ttk.Treeview(frame, columns=columns_CO2)
    rows_CO2 = df_fuel_CO2[columns_CO2].to_numpy().tolist()
    row_index_CO2 = df_fuel_CO2.index.tolist()
    
    table_CO2.heading("#0", text="     Fuel", anchor='w')
    table_CO2.column("#0", stretch=False, width=int(ww*0.11))
    
    def format_value(value):
        if np.isnan(value):
            return ''
        elif round(value,5) == 0.:
            return ''
        else:
            return f'   {value:.3f}'
    
    w = [int(ww*0.072), int(ww*0.072), int(ww*0.052), int(ww*0.057), int(ww*0.048), int(ww*0.036), int(ww*0.036), int(ww*0.044), int(ww*0.041)]
    #insert columns and row content in table
    for i in range(len(columns_CO2)):
        table_CO2.column(columns_CO2[i], width=w[i], stretch=False)
        table_CO2.heading(columns_CO2[i], text=heading_CO2[i], anchor='w')
    for i, row in enumerate(rows_CO2):
        formatted_row = [format_value(value) for value in row]
        table_CO2.insert("", 'end', values=formatted_row, text=row_index_CO2[i])
        
    return table_CO2

def CO2_table_fuel(frame, df_fuel_CO2):
    
    heading_CO2_fuel = ['no CCS', 'w. CCS']
    #Use the totals calculated in the backend (summing all columns here would also add the total columns themselves)
    fuel_CO2 = df_fuel_CO2['Total emissions excl. CCS'].to_numpy()
    fuel_CO2_CCS = df_fuel_CO2['Total emissions incl. CCS'].to_numpy()
    table_CO2_fuel = ttk.Treeview(frame, columns=heading_CO2_fuel)
    rows_CO2_fuel = np.column_stack((fuel_CO2, fuel_CO2_CCS))
    row_index_CO2_fuel = df_fuel_CO2.index.tolist()
    
    table_CO2_fuel.heading("#0", text="     Fuel", anchor='w')
    table_CO2_fuel.column("#0", stretch=False, width=int(ww*0.117))
    
    total_CO2 = sum(rows_CO2_fuel).tolist()
    total_CO2 = ["{:.2f}".format(value) for value in total_CO2]
    table_CO2_fuel.insert("", 'end', values=total_CO2, text='Total emissions') 
    
    total_frame1 = tk.LabelFrame(frame, text='    Total emissions', bg=c_white, fg=c_black, font=('Arial', pt(9), 'bold'), bd=0)
    total_frame2 = tk.LabelFrame(frame, text=total_CO2[0], bg=c_white, fg=c_black, font=('Arial', pt(9), 'bold'), bd=0)
    total_frame3 = tk.LabelFrame(frame, text=total_CO2[1], bg=c_white, fg=c_black, font=('Arial', pt(9), 'bold'), bd=0)
    
    
    for i in range(len(heading_CO2_fuel)):
        table_CO2_fuel.column(heading_CO2_fuel[i], width=int(ww*0.04), stretch=False)
        table_CO2_fuel.heading(heading_CO2_fuel[i], text=heading_CO2_fuel[i], anchor='w')
    for i, row in enumerate(rows_CO2_fuel):
        v = rows_CO2_fuel[i]
        formatted_row = ["{:.2f}".format(value) for value in v]
        table_CO2_fuel.insert("", 'end', values=formatted_row, text=row_index_CO2_fuel[i])   
    
    return table_CO2_fuel, total_frame1, total_frame2, total_frame3
    
    
#----------------------------------------------------------------------------------

#Base year CO2 frame
#Containing 
#----------------------------------------------------------------------------------
CO2_BY_frame = tk.Frame(CO2_frame, width=ww*0.8, height=wh*0.37, bg=c_light)
CO2_BY_frame.place(x=0, y=wh*0.076)

header_CO2_BY = tk.Label(CO2_BY_frame, text="Base year", font=font_subsubheader, bg=c_light, fg=c_dark)
header_CO2_BY.place(x=ww*0.005, y=wh*0.001, anchor='nw')
header_CO2_BYsub = tk.Label(CO2_BY_frame, text="CO2 emissions [mill tons]", font=font_subsubsubheader, bg=c_light, fg=c_mid)
header_CO2_BYsub.place(x=ww*0.08, y=wh*0.005, anchor='nw')

def CO2_table_BY():
    table_CO2_BY = CO2_table(CO2_BY_frame, df_fuel_CO2_BY)
    table_CO2_BY.place(x=ww*0.22, y=wh*0.037, height=wh*0.32)
    #Create a vertical scrollbar
    vsb_BY = ttk.Scrollbar(CO2_BY_frame, orient="vertical", command=table_CO2_BY.yview)
    table_CO2_BY.configure(yscrollcommand=vsb_BY.set)
    vsb_BY.place(x=ww*0.785, y=wh*0.037, height=wh*0.32)
    
    table_CO2_fuel_BY, total_frame1_BY, total_frame2_BY, total_frame3_BY = CO2_table_fuel(CO2_BY_frame, df_fuel_CO2_BY)
    table_CO2_fuel_BY.place(x=ww*0.005, y=wh*0.037, height=wh*0.32)
    
    #Create a vertical scrollbar
    vsb_BY_fuel = ttk.Scrollbar(CO2_BY_frame, orient="vertical", command=table_CO2_fuel_BY.yview)
    table_CO2_fuel_BY.configure(yscrollcommand=vsb_BY_fuel.set)
    vsb_BY_fuel.place(x=ww*0.2, y=wh*0.037, height=wh*0.32)
    
    total_frame1_BY.place(x=ww*0.005, y=wh*0.057, height=wh*0.0213, width=int(ww*0.117))
    total_frame2_BY.place(x=int(ww*0.122), y=wh*0.057, height=wh*0.0213, width=int(ww*0.04))
    total_frame3_BY.place(x=int(ww*0.162), y=wh*0.057, height=wh*0.0213, width=int(ww*0.04))

#----------------------------------------------------------------------------------

#Scenario CO2 frame
#Containing 
#----------------------------------------------------------------------------------
CO2_SC_frame = tk.Frame(CO2_frame, width=ww*0.8, height=wh*0.37, bg=c_light)
CO2_SC_frame.place(x=0, y=wh*0.46)

header_CO2_SC = tk.Label(CO2_SC_frame, text="Scenario", font=font_subsubheader, bg=c_light, fg=c_dark)
header_CO2_SC.place(x=ww*0.005, y=wh*0.001, anchor='nw')
header_CO2_SCsub = tk.Label(CO2_SC_frame, text="CO2 emissions [mill tons]", font=font_subsubsubheader, bg=c_light, fg=c_mid)
header_CO2_SCsub.place(x=ww*0.074, y=wh*0.005, anchor='nw')

def CO2_table_SC():
    table_CO2_SC = CO2_table(CO2_SC_frame, df_fuel_CO2_SC)
    table_CO2_SC.place(x=ww*0.22, y=wh*0.037, height=wh*0.32)
    #Create a vertical scrollbar
    vsb_SC = ttk.Scrollbar(CO2_SC_frame, orient="vertical", command=table_CO2_SC.yview)
    table_CO2_SC.configure(yscrollcommand=vsb_SC.set)
    vsb_SC.place(x=ww*0.785, y=wh*0.037, height=wh*0.32)
    
    table_CO2_fuel_SC, total_frame1_SC, total_frame2_SC, total_frame3_SC = CO2_table_fuel(CO2_SC_frame, df_fuel_CO2_SC)
    table_CO2_fuel_SC.place(x=ww*0.005, y=wh*0.037, height=wh*0.32)
    
    #Create a vertical scrollbar
    vsb_SC_fuel = ttk.Scrollbar(CO2_SC_frame, orient="vertical", command=table_CO2_fuel_SC.yview)
    table_CO2_fuel_SC.configure(yscrollcommand=vsb_SC_fuel.set)
    vsb_SC_fuel.place(x=ww*0.2, y=wh*0.037, height=wh*0.32)
    
    total_frame1_SC.place(x=ww*0.005, y=wh*0.057, height=wh*0.0213, width=int(ww*0.117))
    total_frame2_SC.place(x=int(ww*0.122), y=wh*0.057, height=wh*0.0213, width=int(ww*0.04))
    total_frame3_SC.place(x=int(ww*0.162), y=wh*0.057, height=wh*0.0213, width=int(ww*0.04))
#==============================================================================



#Create "Economy" frame that appears when Economy button is pressed
#==============================================================================
economy_frame = tk.Frame(window)
def show_economy_frame():
    economy_frame.config(width=ww*0.8, height=wh*0.85, bg=c_white) 
    economy_frame.place(x=ww*0.17, y=wh*0.025)
    
header_frame(economy_frame, "Economy", "Base year & Scenario", px(190))

#Creating editable Fuel prices tables
columns_fuelPrices = ['Fuel', 'Price [€/GJ]']
y_pos_prices = wh*0.045
x_pos_prices = ww*0.007

def prices_table(df_FuelPrice, frame, co2_price, year):
    if year == "Scenario year":
        table_fuelPrices = TreeviewEdit(frame, parent_frame=frame, y_pos=y_pos_prices, x_pos=x_pos_prices, columns=columns_fuelPrices, show='headings', unit='', allowed=(0, None), what='Prices', unit_label='€/GJ')
    elif year == "Base year":
        table_fuelPrices = ttk.Treeview(frame, columns=columns_fuelPrices, show='headings')
    rows_fuelPrices = df_FuelPrice[['Market price [€/GJ]']][:].to_numpy().tolist()
    row_index_fuelPrices = df_FuelPrice.index.tolist()
    rows_fuelPrices = np.column_stack((row_index_fuelPrices, rows_fuelPrices))
    rows_editableFuelPrices = []

    # Only include editable fuel prices in table
    for fuel in rows_fuelPrices:
        if df_FuelPrice.loc[fuel[0], 'Editable [y/n?]'] == 'y':
            rows_editableFuelPrices.append(fuel)

    widths = [int(ww*0.11), int(ww*0.05)]
    #Insert columns and row content in table
    for i in columns_fuelPrices:
        table_fuelPrices.column(i, width=widths[columns_fuelPrices.index(i)], stretch=False)
        table_fuelPrices.heading(i, text=columns_fuelPrices[columns_fuelPrices.index(i)], anchor='w')
    for i in range(len(rows_editableFuelPrices)):
        v = rows_editableFuelPrices[i]
        formatted_row = [v[0]] + ["{:.2f}".format(value) for value in v[1:].astype(float)]
        table_fuelPrices.insert("", 'end', values=formatted_row)
      
    table_fuelPrices.place(x=x_pos_prices, y=y_pos_prices)
    
    
    
    

    #Creating button to save electricity table contents to dataframe
    def save_prices_action():
        global co2_price_SC
        df_saved_FuelPrices = save_table_to_dataframe(table_fuelPrices, has_children=False)
        df_FuelPrice.loc[df_FuelPrice.index.isin(df_saved_FuelPrices.index), 'Market price [€/GJ]'] = df_saved_FuelPrices['Price [€/GJ]'].astype(float)
        if year == "Scenario year":
            try:
                new_co2_price = float(entry_co2Price.get().strip().replace(',', '.'))
            except ValueError:
                tk.messagebox.showerror("Invalid input", "Please enter a number for the CO2 price (in €/t CO2).")
                return
            if new_co2_price < 0:
                tk.messagebox.showerror("Invalid input", "The CO2 price must be 0 €/t CO2 or more.\n\nYou entered: " + entry_co2Price.get().strip())
                return
            co2_price_SC = new_co2_price
        
    #save_prices_action()    
    # Edit price of CO2
    label_co2Price = tk.LabelFrame(frame, text="CO2 Price [€/t CO2]",font=font_text_bold, labelanchor='w', bd=0, bg=c_mid, fg=c_white)
    label_co2Price.place(x=x_pos_prices, y=y_pos_prices + wh*0.24, height=wh*0.02, width=ww*0.11)
    if year == "Scenario year":    
        entry_co2Price = ttk.Entry(frame, font=font_table)
        entry_co2Price.insert(0, co2_price)
        entry_co2Price.place(x=x_pos_prices+ww*0.11, y=y_pos_prices + wh*0.24, height=wh*0.019, width=ww*0.05)
        save_prices_button = FlatButton(frame, text="Save price changes", command=save_prices_action, bd=0, bg=c_accent, fg=c_dark, activebackground=c_dark, activeforeground=c_accent, font=font_text_bold)
        save_prices_button.place(x=ww*0.085, y=wh*0.01, height=wh*0.025, width=ww*0.083)
    elif year == "Base year":
        label_co2Price = tk.Label(frame, text=co2_price, font=font_table, bd=0, bg=c_white, fg=c_dark)
        label_co2Price.place(x=x_pos_prices+ww*0.11, y=y_pos_prices + wh*0.24, height=wh*0.019, width=ww*0.05)

    
    
    
    
def costs_table(frame, df_capital_costs, df_om_costs, df_fuel_costs, df_co2_costs):
    heading_costs = ['Electricity', 'District Heat', 'Fuel Refinery', 'Residential', 'Tertiary', 'Industry', 'Agriculture', 'Transport']
    costs_capital = df_capital_costs.sum(axis=0)
    costs_om = df_om_costs.sum(axis=0)
    costs_fuel = df_fuel_costs.sum(axis=0)
    costs_co2 = df_co2_costs.sum(axis=0)
    rows_costs = np.vstack((costs_capital, costs_om, costs_fuel, costs_co2))
    table_costs = ttk.Treeview(frame, columns=heading_costs)
    
   
    row_index_costs = ['Capital costs', 'O&M costs', 'Fuel costs', 'CO2 costs']
    
    table_costs.heading("#0", text="", anchor='w')
    table_costs.column("#0", stretch=False, width=int(ww*0.055))
    
    # total_costs = sum(rows_costs).tolist()
    # total_costs = ["{:.2f}".format(value) for value in total_costs]
    # table_costs.insert("", 'end', values=total_costs, text='Total costs') 
    
    w = [int(ww*0.044), int(ww*0.053), int(ww*0.055), int(ww*0.046), int(ww*0.039), int(ww*0.037), int(ww*0.044), int(ww*0.042)]
    
    for i in range(len(heading_costs)):
        table_costs.column(heading_costs[i], width=w[i], stretch=False)
        table_costs.heading(heading_costs[i], text=heading_costs[i], anchor='w')
    for i in range(len(rows_costs)):
        v = rows_costs[i]
        formatted_row = ["{:.1f}".format(value) for value in v.astype(float)]
        if row_index_costs[i] == 'O&M costs':
            formatted_row[-1] = 'in capital*' #vehicle maintenance is part of the capital costs (see the note next to the table)
        table_costs.insert("", 'end', values=formatted_row, text=row_index_costs[i])

    table_costs.place(x=x_pos_prices + ww*0.17, y=y_pos_prices, height = wh*0.115)

    note_costs = tk.Label(frame, text="* Transport: the capital costs are the costs of the vehicles per person km or tonne km, investment and maintenance together (as in the Excel STREAM model).", font=font_text, bg=c_light, fg=c_dark, justify='left', anchor='nw', wraplength=int(ww*0.17))
    note_costs.place(x=x_pos_prices + ww*0.6, y=y_pos_prices, width=ww*0.18, height=wh*0.115)
    

def costs_plot(frame, df_capital_costs, df_om_costs, df_fuel_costs, df_co2_costs, year):
    heading_costs = ['Electricity', 'District Heat', 'Fuel Refinery', 'Residential', 'Tertiary', 'Industry', 'Agriculture', 'Transport']
    costs_capital = df_capital_costs.sum(axis=0)
    costs_om = df_om_costs.sum(axis=0)
    costs_fuel = df_fuel_costs.sum(axis=0)
    costs_co2 = df_co2_costs.sum(axis=0)
    rows_costs = pd.DataFrame(np.vstack((costs_capital, costs_om, costs_fuel, costs_co2)), columns=heading_costs, index=['Capital costs', 'O&M costs', 'Fuel costs', 'CO2 costs'])
    fig, ax = plot_stacked_bar(rows_costs, '', len(heading_costs), (10, 2))
    fig.suptitle("Yearly costs - " + str(year), fontsize=10)
    ax.set_ylabel("[mill. €]", fontsize=8)
    ax.set_xticklabels(heading_costs,rotation=20, fontsize=7)
    canvas_costs = FigureCanvasTkAgg(fig, master=frame)
    canvas_costs.get_tk_widget().place(x=x_pos_prices + ww*0.17, y=wh*0.17, width=ww*0.54, height=wh*0.19)
    imageName = figure_path("costs" + str(year) + ".png")
    fig.savefig(imageName)

    return fig, ax

#Base year Economy frame
#Containing 
#----------------------------------------------------------------------------------
economy_BY_frame = tk.Frame(economy_frame, width=ww*0.8, height=wh*0.37, bg=c_light)
economy_BY_frame.place(x=0, y=wh*0.076)

header_economy_BY = tk.Label(economy_BY_frame, text="Base year", font=font_subsubheader, bg=c_light, fg=c_dark)
header_economy_BY.place(x=ww*0.005, y=wh*0.001, anchor='nw')


def prices_BY():
    
    prices_table(df_FuelPrice_BY, economy_BY_frame, co2_price_BY, "Base year")
    
    #missing subheader with 'Costs [mill. €]'
    prices_subheader = tk.Label(economy_BY_frame, text="Costs [mill. €]", font=font_subsubsubheader, bg=c_light, fg=c_mid)
    prices_subheader.place(x=x_pos_prices + ww*0.17, y=wh*0.007, anchor='nw')
    
    costs_table(economy_BY_frame, df_capital_costs_BY, df_om_costs_BY, df_fuel_costs_BY, df_co2_costs_BY)
    costs_plot(economy_BY_frame, df_capital_costs_BY, df_om_costs_BY, df_fuel_costs_BY, df_co2_costs_BY, Base_Year)
    
prices_BY()

#----------------------------------------------------------------------------------

#Scenario Economy frame
#Containing 
#----------------------------------------------------------------------------------
economy_SC_frame = tk.Frame(economy_frame, width=ww*0.8, height=wh*0.37, bg=c_light)
economy_SC_frame.place(x=0, y=wh*0.46)

header_economy_SC = tk.Label(economy_SC_frame, text="Scenario", font=font_subsubheader, bg=c_light, fg=c_dark)
header_economy_SC.place(x=ww*0.005, y=wh*0.001, anchor='nw')

def prices_SC():
    prices_table(df_FuelPrice_SC, economy_SC_frame, co2_price_SC, "Scenario year")
    #missing subheader with 'Costs [mill. €]'
    prices_subheader = tk.Label(economy_SC_frame, text="Costs [mill. €]", font=font_subsubsubheader, bg=c_light, fg=c_mid)
    prices_subheader.place(x=x_pos_prices + ww*0.17, y=wh*0.007, anchor='nw')
    costs_table(economy_SC_frame, df_capital_costs_SC, df_om_costs_SC, df_fuel_costs_SC, df_co2_costs_SC)
    costs_plot(economy_SC_frame, df_capital_costs_SC, df_om_costs_SC, df_fuel_costs_SC, df_co2_costs_SC, Scenario_Year)
    
    
prices_SC()


#==============================================================================


#Create "Time Series" frame that appears when Time Series button is pressed
#==============================================================================
timeSeries_frame = tk.Frame(window)
timeSeries_subframe = tk.Frame(timeSeries_frame, width=ww*0.8, height=wh*0.8, bg=c_light)
#The time series takes a while to calculate, so it is calculated when its page is opened (again after the scenario changes)
df_hour, technologies_used = None, None
header_frame(timeSeries_frame, "Time Series", "Scenario", px(250))

#Show time series frame
def show_timeSeries_frame():
    timeSeries_frame.config(width=ww*0.8, height=wh*0.85, bg=c_white) 
    timeSeries_frame.place(x=ww*0.17, y=wh*0.025)
    timeSeries_subframe.place(x=0, y=wh*0.076)
    if df_hour is None: #not calculated yet for this scenario
        run_while_busy(update_timeSeries, "Calculating the time series...")
    else:
        durationCurve_plot(df_hour['Consumption after renewables (unflexible and flexible)'], technologies_used)
        timeSeries_avgWeek_plot()

#Plot duration curve
def durationCurve_plot(consumption_afterRenew, technologies_used):
    sorted_consumption_afterRenew = consumption_afterRenew.sort_values(ascending=False).to_numpy()
    fig, ax = plot_duration_curve(sorted_consumption_afterRenew, technologies_used)
    canvas_durationCurve = FigureCanvasTkAgg(fig, master=timeSeries_subframe)
    canvas_durationCurve.get_tk_widget().place(x=ww*0.005, y=wh*0.004, width = ww*0.5, height = wh*0.4)
    imageName = figure_path("durationCurve" + ".png")
    fig.savefig(imageName)  

#Plot power production per technology average week
def timeSeries_avgWeek_plot():
    fig, ax = plot_timeSeries_avgWeek(df_hour)
    canvas_avgWeek = FigureCanvasTkAgg(fig, master=timeSeries_subframe)
    canvas_avgWeek.get_tk_widget().place(x=-ww*0.034, y=wh*0.36, width = ww*0.8, height = wh*0.4)
    imageName = figure_path("avgWeek" + ".png")
    fig.savefig(imageName)

def timeSeries_avgMonth_plot():
    fig, ax = plot_timeSeries_avgMonth(df_hour)
    canvas_avgMonth = FigureCanvasTkAgg(fig, master=timeSeries_subframe)
    canvas_avgMonth.get_tk_widget().place(x=-ww*0.034, y=wh*0.36, width = ww*0.8, height = wh*0.4)
    imageName = figure_path("avgMonth" + ".png")
    fig.savefig(imageName)

button_plot_avgWeek = FlatButton(timeSeries_subframe, text="Average week", command=timeSeries_avgWeek_plot, bd=0, bg=c_midlight, fg=c_white, activebackground=c_dark, activeforeground=c_accent, font=font_text_bold)
button_plot_avgWeek.place(x=ww*0.55, y=wh*0.33, height=wh*0.032, width=ww*0.1)
button_plot_avgMonth = FlatButton(timeSeries_subframe, text="Average month", command=timeSeries_avgMonth_plot, bd=0, bg=c_midlight, fg=c_white, activebackground=c_dark, activeforeground=c_accent, font=font_text_bold)
button_plot_avgMonth.place(x=ww*0.66, y=wh*0.33, height=wh*0.032, width=ww*0.1)

#Update time series data
def update_timeSeries():
    global df_hour, technologies_used
    df_hour, technologies_used = run_timeSeries(df_Power_SC, df_FuelPrice_SC, df_FEC_SC, total_electricity_demand_SC, el_transport_SC, el_DH_SC, el_FR_SC)
    durationCurve_plot(df_hour['Consumption after renewables (unflexible and flexible)'], technologies_used)
    timeSeries_avgWeek_plot()

#Button to update time series data
button_update_timeSeries = FlatButton(timeSeries_subframe, text='Update Time Series', command=lambda: run_while_busy(update_timeSeries, "Calculating the time series..."), bd=0, bg=c_dark, fg=c_accent, activebackground=c_accent, activeforeground=c_dark, font=font_text_bold)
button_update_timeSeries.place(x=ww*0.69, y=wh*0.01, height=wh*0.035, width=ww*0.1)
text_update_timeSeries = tk.Label(timeSeries_subframe, text="May take a minute", font=font_text, bg=c_light, fg=c_dark)
text_update_timeSeries.place(x=ww*0.69, y=wh*0.045, anchor='nw')


#Check that shares add up before updating
#==============================================================================
def shares_add_up():
    """Return True if all shares that must add up to 100 % do. If not, explain which ones and return False."""
    problems = []
    def check(name, total):
        if abs(total - 1) > 1e-6:
            problems.append(name + ": " + "{:.2%}".format(total))
    for sector in SECTORS:
        check("Heating, " + sector, df_beta_SC[sector].iloc[1:].sum())
    for transport_type in df_TransportFuel_SC.index:
        if transport_type != 'Bike': #bikes use no fuel
            check("Transport fuels, " + transport_type, df_TransportFuel_SC.loc[transport_type].sum())
    allocation = df_TransportWork_SC['Allocation of transport work']
    check("Allocation of transport work, person", allocation.iloc[:5].sum())
    check("Allocation of transport work, goods", allocation.iloc[5:].sum())
    if problems:
        tk.messagebox.showerror("Shares do not add up to 100 %",
            "The scenario was not updated, because these shares must add up to 100 %, but do not:\n\n" + "\n".join(problems) +
            "\n\nCorrect them, press the yellow save button below the table, and update again.")
    return not problems

#Update scenario year
#==============================================================================
def update_SC_year():
    max_year = 2100
    try:
       sc_year = int(entry_year.get())
    except ValueError:
       tk.messagebox.showerror("Invalid Input", "Please enter a valid integer")
       return
    if not Base_Year < sc_year <= max_year:
       #Handle invalid input, input can only be a number between base year and max year
       tk.messagebox.showerror("Invalid Input", ("Please enter a valid year between " + str(Base_Year) + " and " + str(max_year) + "."))
       return
    if not shares_add_up():
        return
    global Scenario_Year
    Scenario_Year = sc_year

    file_path = SCData
    workbook = openpyxl.load_workbook(file_path)
    sheet_year = workbook['Scenario Year']
    sheet_year.cell(row=2, column=1, value=Scenario_Year)
    workbook.save(file_path)
    
    global SC_NEC_tFE_el, SC_NEC_tFE_heat, df_beta_BY, engr 
    SC_NEC_tFE_el, SC_NEC_tFE_heat, df_beta_BY, engr = BY_calculations()
    set_SC_transport_demand(df_Transport_SC, engr)

    update_scenario()
    
entry_year = ttk.Entry(window, font=font_table)
entry_year.insert(0, Scenario_Year)
entry_year.place(x=ww*0.755, y=wh*0.01, height=wh*0.03, width=ww*0.1)
button_update_scenarioYear = FlatButton(window, text='Update Scenario Year', command=lambda: run_while_busy(update_SC_year, "Updating the scenario year..."), bd=0, bg=c_mid, fg=c_white, activebackground=c_accent, activeforeground=c_white, font=font_table_heading)
button_update_scenarioYear.place(x=ww*0.755, y=wh*0.04769, height=wh*0.032, width=ww*0.1)
    

#Update Scenario year data
#==============================================================================
def update_scenario():
    if not shares_add_up():
        return
    save_scenario_file(SCData)
    update_calculations()

#Save the scenario data that can be changed in pySTREAM into a scenario workbook
def save_scenario_file(file_path):
    workbook = openpyxl.load_workbook(file_path)
    #Update power production shares
    power_shares = df_Power_SC['Shares'].tolist()
    sheet_power = workbook['Electricity Production'] 
    r = 2
    for value in power_shares:
        sheet_power.cell(row=r, column=2, value=value)
        r += 1
        
    #Update district heat production shares
    DH_shares = df_Heat_SC['Shares'].tolist()
    sheet_DH = workbook['District Heat Boiler Production']
    r = 2
    for value in DH_shares:
        sheet_DH.cell(row=r, column=2, value=value)
        r += 1
    
    #Update heat sectors consumption shares
    heatSectors_shares = df_beta_SC.to_numpy().tolist()
    sheet_heatSectors = workbook['Final Energy Consumption']
    c = 2
    for i in range(df_beta_SC.shape[1]):
        r = 2
        for j in range(df_beta_SC.shape[0]):
            sheet_heatSectors.cell(row=r, column=c, value=heatSectors_shares[j][i])
            r +=1
        c += 1
    
    #Update transport fuels consumption shares
    transportFuels_shares = df_TransportFuel_SC.to_numpy().tolist()
    sheet_transportFuels = workbook['Fuels Transportation']
    c = 2
    for i in range(df_TransportFuel_SC.shape[1]):
        r = 2
        for j in range(df_TransportFuel_SC.shape[0]):
            sheet_transportFuels.cell(row=r, column=c, value=transportFuels_shares[j][i])
            r +=1
        c += 1
    
    #Update transport allocation and utilisation shares
    transportWork_shares = df_TransportWork_SC.to_numpy().tolist()
    sheet_transportWork = workbook['Transport Work']
    c = 2
    for i in range(df_TransportWork_SC.shape[1]):
        r = 2
        for j in range(df_TransportWork_SC.shape[0]):
            sheet_transportWork.cell(row=r, column=c, value=transportWork_shares[j][i])
            r +=1
        c += 1

    #Update fuel prices
    fuel_prices = df_FuelPrice_SC['Market price [€/GJ]'].tolist()
    sheet_fuelPrices = workbook['Fuel Price']
    r = 2
    for value in fuel_prices:
        sheet_fuelPrices.cell(row=r, column=2, value=value)
        r += 1

    #Update CO2 price
    sheet_co2Price = workbook['CO2 Price']
    sheet_co2Price.cell(row=2, column=1, value=co2_price_SC)
    
    workbook.save(file_path)
    
def update_calculations():
    plt.close() #delete old plots to save memory
    #Calculate all variables for scenario year with updated data
    
    #Variables are made global since they cannot be returned because the function was called with a button
    global df_FEC_SC, df_transport_consumption_SC, df_transport_consumption_fuels_SC, df_fuel_secondary_consumption_SC, total_DH_demand_SC, total_electricity_consumption_SC, gridloss_SC, Fc_DH_SC, Fc_El_SC, total_fuel_consumption_SC, DH_boiler_fuel_consumption_SC, df_Power_SC, df_Fc_p_SC, df_Fc_SC, df_fuelType_consumption_SC, df_fuel_CO2_SC, df_fuelType_CO2_SC, df_fuelCategory_consumption_SC, df_capital_costs_SC, df_om_costs_SC, df_co2_costs_SC, df_FuelPrice_SC, df_fuel_costs_SC, total_electricity_demand_SC, el_transport_SC, el_DH_SC, el_FR_SC, co2_price_SC
    SC_FEC_ES, SC_FEC_tFE, SC_FEC_tES = SC_calculations(SC_NEC_tFE_el, SC_NEC_tFE_heat)
    df_FEC_SC, df_FEC_SC_total = FEC_dataframe(SC_FEC_ES, SC_FEC_tFE, SC_FEC_tES) 
    
    df_transport_consumption_SC, df_transport_consumption_fuels_SC = calculate_transport('Scenario Year', 
                                                                                         df_FEC_SC, 
                                                                                         df_Transport_SC, 
                                                                                         df_TransportWork_SC, 
                                                                                         df_TransportFuel_SC
                                                                                         )  
    df_fuel_secondary_consumption_SC = calculate_FR_fuel(df_FEC_SC, df_transport_consumption_fuels_SC, df_FuelRefinery_SC)
    total_DH_demand_SC = calculate_heat_demand(df_FEC_SC, df_fuel_secondary_consumption_SC, df_FuelRefinery_SC)  
    total_electricity_consumption_SC, gridloss_SC, df_Heat_SC['Amount'], total_electricity_demand_SC, el_transport_SC, el_DH_SC, el_FR_SC  = final_heatPower_calculate(df_FEC_SC, df_transport_consumption_fuels_SC, df_fuel_secondary_consumption_SC, total_DH_demand_SC, df_Power_SC)
    Fc_DH_SC, Fc_El_SC, total_fuel_consumption_SC, DH_boiler_fuel_consumption_SC, df_Power_SC['Total fuel consumption'] = fuel_composition(df_Power_SC['Fuel'], 
                                                                                                                                           df_Power_SC['Electricity production (condensing or nonCHP)'], 
                                                                                                                                           df_Power_SC['Electric efficiency (extract unit: condensing mode)'], 
                                                                                                                                           df_Power_SC['Fuel consumption in CHP-mode'], 
                                                                                                                                           df_Power_SC['CHP heat production'], 
                                                                                                                                           df_Power_SC['Heat efficiency in CHP mode (used for allocation)'], 
                                                                                                                                           df_Power_SC['Mode'], 
                                                                                                                                           df_Heat_SC['Fuel'], 
                                                                                                                                           df_Heat_SC['Amount'], 
                                                                                                                                           df_Heat_SC['Fuel efficiency']
                                                                                                                                           )
    df_Fc_p_SC, df_Fc_SC, df_fuelType_consumption_SC, df_fuelCategory_consumption_SC = resources(df_FEC_SC,  
                                                                 df_transport_consumption_fuels_SC, 
                                                                 Fc_El_SC, 
                                                                 Fc_DH_SC, 
                                                                 total_fuel_consumption_SC, 
                                                                 df_FuelRefinery_SC, 
                                                                 DH_boiler_fuel_consumption_SC, 
                                                                 df_fuel_secondary_consumption_SC, 
                                                                 gridloss_SC, 
                                                                 total_electricity_consumption_SC, 
                                                                 df_Power_SC['CHP heat production'].sum(), 
                                                                 df_Heat_SC['Amount'].sum()
                                                                 ) 
   
    df_fuel_CO2_SC, df_fuelType_CO2_SC = calculate_CO2(df_Fc_SC, df_Power_SC, df_TechEff_SC)

    df_capital_costs_SC = capital_costs(df_Power_SC, df_Heat_SC, df_Fc_SC, df_FEC_SC, df_SectorsDetails_SC, df_FuelRefineryCost_SC, df_fuel_secondary_consumption_SC,
                                        transport_costs(df_Transport_SC, df_TransportWork_SC, df_TransportFuel_SC, df_TransportCosts_SC))
    df_om_costs_SC = om_costs(df_Power_SC, df_Heat_SC, df_Fc_SC, df_FEC_SC, df_SectorsDetails_SC, df_FuelRefineryCost_SC, df_fuel_secondary_consumption_SC)
    df_co2_costs_SC = co2_costs(co2_price_SC, df_fuel_CO2_SC)
    df_FuelPrice_SC, df_fuel_costs_SC = fuel_costs(df_FuelPrice_SC, df_FuelRefinery_SC, df_FuelRefineryCost_SC, df_Fc_SC)


    sankey_SC()
    sankey_BY()
    heatConsumption_table()
    heat_sums()
    transportFuels_table()
    transportFuels_sums()
    allocation_table()
    transportAllocation_sums()
    utilisation_table()
    electricity_table()
    heat_table()
    electricity_amounts()
    DH_amounts()
    resources_table()
    electricity_sums()
    DH_sums()
    resources_plot()
    resources_table()
    plots_production()
    create_FEC_SC_frame()
    create_production_SC_frame()
    transport_table_amounts_SC()
    total_fuel_transport_table_SC()
    transport_table_shares_SC()
    transport_table_amounts_SC()
    total_fuel_transport_table_SC()
    CO2_table_SC()
    prices_SC()


    header_frame(scenario_frame, "Scenario Year Data", str(Scenario_Year), ww*0.2)

    #The time series belongs to the previous scenario now; calculate it again when its page is shown
    global df_hour, technologies_used
    df_hour, technologies_used = None, None
    if timeSeries_frame.winfo_ismapped():
        show_timeSeries_frame()

    #Warn when the CHP plants alone produce more district heat than is needed; the surplus is not used
    surplus_heat = df_Power_SC['CHP heat production'].sum() - total_DH_demand_SC
    if surplus_heat > 0:
        tk.messagebox.showwarning("Surplus district heat",
            "The combined heat and power (CHP) plants produce {:.1f} PJ more district heat than is needed ({:.1f} PJ). "
            "No district heating boilers are used and the surplus heat is not used.\n\n"
            "To avoid this, produce less electricity in CHP plants or use more district heating.".format(surplus_heat, total_DH_demand_SC))


#Scenario files: save as, open, reset to default
#==============================================================================
SCENARIO_SHEETS = ['Scenario Year', 'Final Energy Consumption', 'Electricity Production', 'District Heat Boiler Production', 'Transport',
                   'Fuels Transportation', 'Transport Work', 'Technology Improvements', 'Energy Savings', 'Technology Efficiencies',
                   'Fuel Refinery Fuel Input Factor', 'Fuel Price', 'Fuel Refinery Cost', 'Sectors Details', 'CO2 Price']
#the scenario data in the order load_SCdata() returns it
SCENARIO_DATA_NAMES = ['Scenario_Year', 'df_beta_SC', 'df_Power_SC', 'df_Heat_SC', 'df_Transport_SC', 'df_TransportFuel_SC', 'df_TransportWork_SC',
                       'df_TransportImprovement_SC', 'df_EnergySavings_SC', 'df_TechEff_SC', 'df_FuelRefinery_SC', 'df_FuelPrice_SC',
                       'df_FuelRefineryCost_SC', 'df_SectorsDetails_SC', 'co2_price_SC', 'df_TransportCosts_SC']

def save_scenario_as(file_path=None):
    """'Save scenario as': save the current scenario in a file chosen by the user (file_path is used by the tests)."""
    if not shares_add_up():
        return
    if file_path is None:
        file_path = filedialog.asksaveasfilename(title="Save scenario as", defaultextension=".xlsx",
                                                 filetypes=[("Excel workbook", "*.xlsx")], initialfile="My scenario.xlsx")
        if not file_path: #cancelled
            return
    save_scenario_file(SCData)
    shutil.copyfile(SCData, file_path)
    tk.messagebox.showinfo("Scenario saved", "The scenario was saved as:\n" + file_path + "\n\nUse 'Open scenario' to continue with it later.")

def open_scenario(file_path=None):
    """'Open scenario': continue with a scenario saved earlier (file_path is used by the tests)."""
    if file_path is None:
        file_path = filedialog.askopenfilename(title="Open scenario", filetypes=[("Excel workbook", "*.xlsx")])
        if not file_path: #cancelled
            return
    load_scenario_file(file_path)

def reset_scenario():
    """'Reset to default scenario': go back to the scenario that comes with pySTREAM."""
    if tk.messagebox.askyesno("Reset to default scenario", "This replaces the current scenario with the default scenario that comes with pySTREAM.\n\n"
                              "To keep the current scenario, choose No and use 'Save scenario as' first.\n\nReset now?"):
        load_scenario_file(DEFAULT_SCData)

def load_scenario_file(file_path):
    """Make file_path the current scenario (copied to data/Scenario Data.xlsx) and recalculate everything."""
    with pd.ExcelFile(file_path) as workbook:
        missing = [sheet for sheet in SCENARIO_SHEETS if sheet not in workbook.sheet_names]
    if missing:
        tk.messagebox.showerror("Not a scenario file", "This file is not a pySTREAM scenario (it has no sheet '" + missing[0] + "'):\n" + file_path)
        return
    with open(SCData, 'rb') as current_file:
        previous_scenario = current_file.read()
    if os.path.abspath(file_path) != os.path.abspath(SCData):
        shutil.copyfile(file_path, SCData)
    try:
        reload_scenario_data()
    except Exception:
        #put the previous scenario back, then show the error
        with open(SCData, 'wb') as current_file:
            current_file.write(previous_scenario)
        reload_scenario_data()
        raise
    run_while_busy(update_calculations, "Calculating the scenario...")

def reload_scenario_data():
    """Read the scenario data file again, both here and in the backend (whose functions use their own copies)."""
    import Backend
    global SC_NEC_tFE_el, SC_NEC_tFE_heat, df_beta_BY, engr
    SC_NEC_tFE_el, SC_NEC_tFE_heat, df_beta_BY, engr = BY_calculations() #these depend on the scenario year
    Backend.engr = engr #load_SCdata() uses it for the transport demand
    for name, value in zip(SCENARIO_DATA_NAMES, load_SCdata()):
        setattr(Backend, name, value)
        globals()[name] = value
    entry_year.delete(0, 'end')
    entry_year.insert(0, Scenario_Year)

scenario_file_label = tk.Label(menu_frame, text="Scenario file", font=font_text_bold, bg=c_mid, fg=c_white)
scenario_file_label.place(x=ww*0.014, y=wh*0.795)
for i, (text, command) in enumerate([("Save scenario as...", save_scenario_as), ("Open scenario...", open_scenario), ("Reset to default scenario", reset_scenario)]):
    scenario_file_button = FlatButton(menu_frame, text=text, command=command, bd=0, bg=c_dark, fg=c_white, activebackground=c_accent, activeforeground=c_dark, font=font_text_bold)
    scenario_file_button.place(x=ww*0.014, y=wh*(0.82 + 0.035*i), width=button_width, height=wh*0.03)

button_update_scenario = FlatButton(window, text='Update Scenario', command=lambda: run_while_busy(update_scenario, "Updating the scenario..."), bd=0, bg=c_dark, fg=c_white, activebackground=c_accent, activeforeground=c_white, font=font_button)

button_update_scenario.place(x=ww*0.86, y=wh*0.01, height=wh*0.07, width=ww*0.1)
#==============================================================================

show_graphs_frame()
starting_label.destroy()
#show_timeSeries_frame()
#show_economy_frame()


root.mainloop()
