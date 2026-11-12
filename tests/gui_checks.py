"""Drive the pySTREAM GUI from a script, for the GUI tests (see test_gui.py).

Loads pySTREAM.py from a copy of the program without starting the main loop
(so no window waits for a user), runs one named check and writes what it
observed to a JSON file. Message boxes are recorded instead of shown.

Usage: python gui_checks.py <program_dir> <check> <output.json>
"""
import faulthandler
import json
import os
import sys
import tempfile
import tkinter
import tkinter.messagebox
import tkinter.ttk
import warnings

import helpers

# If a check hangs, print where every thread is and stop, so the test shows where it got stuck
faulthandler.dump_traceback_later(240, exit=True)

program_dir, check, output_file = sys.argv[1:4]
warnings.simplefilter('ignore')
# Normally the GUI is started from the program folder; this check starts it from somewhere else
start_folder = tempfile.mkdtemp() if check == 'start_from_other_folder' else program_dir
os.chdir(start_folder)
sys.path.insert(0, program_dir)

messages = []
for kind in ('showerror', 'showinfo', 'showwarning'):
    setattr(tkinter.messagebox, kind, lambda title, message, _kind=kind, **kw: messages.append([_kind, title, message]))
# questions are answered 'yes'
tkinter.messagebox.askyesno = lambda title, message, **kw: messages.append(['askyesno', title, message]) or True
tkinter.Tk.mainloop = lambda self, n=0: None

# Kaleido (which renders the Sankey diagrams to PNG with Chrome) sometimes hangs for minutes,
# so the checks save a small placeholder image instead. Everything else runs as normal.
# Two checks instead simulate Kaleido failing (e.g. no Chrome installed) or hanging.
import time
import plotly.graph_objects
from PIL import Image

def write_placeholder(self, file, *args, **kwargs):
    Image.new('RGB', (40, 30), 'white').save(file)

def write_fails(self, file, *args, **kwargs):
    raise RuntimeError('Kaleido requires Google Chrome to be installed')

def write_hangs(self, file, *args, **kwargs):
    time.sleep(3600)

plotly.graph_objects.Figure.write_image = {'sankey_image_fails': write_fails,
                                           'sankey_image_hangs': write_hangs}.get(check, write_placeholder)

def read_error_log():
    path = os.path.join(program_dir, 'pySTREAM error log.txt')
    return open(path, encoding='utf-8').read() if os.path.exists(path) else None


if check == 'missing_data_file':
    # Start without one of the data files; pySTREAM must explain the problem instead of only printing a traceback
    os.remove(os.path.join(program_dir, 'data', 'Assumptions.xlsx'))
    try:
        import pySTREAM as gui
    except Exception:
        sys.excepthook(*sys.exc_info())  # what Python does with an error that stops the program
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump({'messages': messages, 'log': read_error_log()}, f)
    os._exit(0)

# The screen_* checks pretend the screen has another size. The window then cannot fill the screen ('zoomed'),
# so pySTREAM uses the screen size minus 100 pixels for the taskbar and title bar: the layout gets width x height.
SIMULATED_SCREENS = {'screen_laptop': (1366, 768), 'screen_design': (1920, 1080),
                     'screen_large': (2560, 1440), 'screen_tiny': (1024, 600)}
if check in SIMULATED_SCREENS:
    width, height = SIMULATED_SCREENS[check]
    tkinter.Misc.winfo_screenwidth = lambda self: width
    tkinter.Misc.winfo_screenheight = lambda self: height + 100
    tkinter.Wm.wm_state = tkinter.Wm.state = lambda self, newstate=None: 'normal' if newstate is None else None

import pySTREAM as gui
gui.root.withdraw()


def co2_totals():
    """Totals shown on the CO2 page next to the model's own totals."""
    out = {}
    for year, df in [('BY', gui.df_fuel_CO2_BY), ('SC', gui.df_fuel_CO2_SC)]:
        table, _, total_no_ccs, total_with_ccs = gui.CO2_table_fuel(gui.CO2_SC_frame, df)
        rows = {table.item(item, 'text'): [float(v) for v in table.item(item, 'values')] for item in table.get_children()}
        out[year] = {
            'shown_total': [float(total_no_ccs.cget('text')), float(total_with_ccs.cget('text'))],
            'model_total': [df['Total emissions excl. CCS'].sum(), df['Total emissions incl. CCS'].sum()],
            'shown_rows': {fuel: values for fuel, values in rows.items() if fuel != 'Total emissions'},
            'model_rows': {fuel: [df.loc[fuel, 'Total emissions excl. CCS'], df.loc[fuel, 'Total emissions incl. CCS']] for fuel in df.index},
        }
    return out


def update_matches_fresh_start():
    """Change the scenario the way the edit tables do, press 'Update Scenario', and
    compare with a fresh start of the model on the data file the GUI just saved."""
    # 'Edit production': move 10 %-points of electricity from coal with CO2 storage to offshore wind
    gui.df_Power_SC.loc['Coal incl. CO2-storage', 'Shares'] -= 0.10
    gui.df_Power_SC.loc['Wind, offshore', 'Shares'] += 0.10
    # 'Edit consumption': move 10 %-points of residential heating from district heating to heat pumps
    gui.df_beta_SC.loc['Heatpumps', 'Residential'] += 0.10
    gui.df_beta_SC.loc['District Heating', 'Residential'] -= 0.10
    gui.update_scenario()
    gui.update_timeSeries()
    after_update = helpers.collect_outputs(gui)
    after_update['time_series'] = helpers.collect_time_series(gui.df_hour, gui.technologies_used)
    return {'after_update': after_update,
            'fresh_start': helpers.run_model(program_dir, time_series=True)}


def update_time_series_button():
    """Press 'Update Time Series' and report any error."""
    try:
        gui.update_timeSeries()
        return {'exception': None}
    except Exception as e:
        return {'exception': f'{type(e).__name__}: {e}'}


def update_scenario_year():
    """Enter invalid years (nothing may be saved), then change the year to 2035 and
    compare with a fresh start of the model on the saved data."""
    data_file = os.path.join(program_dir, 'data', 'Scenario Data.xlsx')
    file_before = open(data_file, 'rb').read()
    exceptions = []

    def enter_year(text):
        gui.entry_year.delete(0, 'end')
        gui.entry_year.insert(0, text)
        try:
            gui.update_SC_year()
        except Exception as e:
            exceptions.append(f'{text}: {type(e).__name__}: {e}')

    for text in ['abc', '1990', '2200']:
        enter_year(text)
    invalid_input = {'messages': list(messages), 'file_unchanged': open(data_file, 'rb').read() == file_before}
    enter_year('2035')
    return {'invalid_input': invalid_input, 'exceptions': exceptions,
            'after_update': helpers.collect_outputs(gui), 'fresh_start': helpers.run_model(program_dir)}


def start_from_other_folder():
    """The GUI was started from another folder (see top of file). Press 'Update Scenario'
    and check that the data file in the program folder was saved and nothing was written elsewhere."""
    data_file = os.path.join(program_dir, 'data', 'Scenario Data.xlsx')
    modified_before = os.path.getmtime(data_file)
    gui.update_scenario()
    return {'data_file_saved': os.path.getmtime(data_file) > modified_before,
            'files_in_start_folder': os.listdir(start_folder),
            'figures_folder': sorted(os.listdir(os.path.join(program_dir, 'figures')))
                              if os.path.isdir(os.path.join(program_dir, 'figures')) else None}


def chp_surplus_warning():
    """Choose so much coal CHP that CHP heat exceeds the district heating demand, then press 'Update Scenario'."""
    gui.df_Power_SC.loc['Coal incl. CO2-storage', 'Shares'] += 0.40
    gui.df_Power_SC.loc['Wind, onshore', 'Shares'] -= 0.244
    gui.df_Power_SC.loc['Wind, offshore', 'Shares'] -= 0.156
    gui.update_scenario()
    return {}


def sankey_buttons():
    """Texts of the buttons in the Graphs page's two diagram areas (a button with an image has no text)."""
    frames = [gui.graphs_sc_frame, gui.graphs_by_frame]
    return {'button_texts': [w.cget('text') for frame in frames for w in frame.winfo_children() if isinstance(w, gui.FlatButton)]}


def close_window():
    """Close the window like a user would; pySTREAM must then stop right away."""
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump({'stopped_when_closed': True, 'messages': []}, f)
    gui.quit_me()
    return {'stopped_when_closed': False}


def economy_page():
    """What the cost tables on the Economy page show, the notes next to them, and the model's transport capital costs."""
    out = {}
    for year, frame, capital in [('BY', gui.economy_BY_frame, gui.df_capital_costs_BY), ('SC', gui.economy_SC_frame, gui.df_capital_costs_SC)]:
        out[year + ' transport capital costs'] = float(capital['Transportation'].sum())
        tables = [w for w in frame.winfo_children() if isinstance(w, tkinter.ttk.Treeview) and 'Transport' in w['columns']]
        out[year] = {tables[-1].item(i, 'text'): list(tables[-1].item(i, 'values')) for i in tables[-1].get_children()}
        out[year + ' notes'] = [str(w.cget('text')) for w in frame.winfo_children() if isinstance(w, tkinter.Label)]
    return out


def plots():
    """What three plots show: Sankey flows from primary sources, the duration curve's y-axis
    range, and the labels on resource bars that are cut off."""
    import matplotlib.pyplot as plt
    out = {}
    for year, ft, fc in [('BY', gui.df_fuelType_consumption_BY, gui.df_Fc_BY), ('SC', gui.df_fuelType_consumption_SC, gui.df_Fc_SC)]:
        diagram = gui.sankey_diagram(ft, fc, None)
        out['sankey ' + year] = {'shown': sum(v for s, v in zip(diagram.link.source, diagram.link.value) if s <= 10),
                                 'model': ft['Total primary fuel consumption'].drop(['Electricity', 'District Heating']).sum()}
    # more wind gives deeper surplus hours, i.e. more negative values on the duration curve
    gui.df_Power_SC.loc['Coal incl. CO2-storage', 'Shares'] -= 0.20
    gui.df_Power_SC.loc['Wind, offshore', 'Shares'] += 0.20
    gui.update_scenario()
    gui.update_timeSeries()
    curve = gui.df_hour['Consumption after renewables (unflexible and flexible)'].sort_values(ascending=False).to_numpy()
    fig, ax = gui.plot_duration_curve(curve, gui.technologies_used)
    out['duration curve'] = {'curve': [curve.min(), curve.max()], 'axis': list(ax.get_ylim())}
    gui.resources_plot()
    resources = gui.fix_resources_rows()
    out['resources'] = {'longest bars': sorted(float(v) for v in resources[:, 1:].ravel() if float(v) > 500),
                        'labels': [t.get_text() for a in plt.gcf().axes for t in a.texts]}
    return out


def file_open_in_excel():
    """Press 'Update Scenario' while the data file is open in Excel (which locks it)."""
    import openpyxl
    def locked(path, *args, **kwargs):
        raise PermissionError(13, 'Permission denied', str(path))
    openpyxl.load_workbook = locked
    gui.button_update_scenario.invoke()
    return {'log': read_error_log()}


def input_validation():
    """Type values into the editable tables like a user (double-click a cell, type, press Enter),
    and press 'Update Scenario' with shares that do not add up to 100 %."""
    from types import SimpleNamespace

    def last_table(frame):
        return [w for w in frame.winfo_children() if isinstance(w, gui.TreeviewEdit)][-1]

    def enter(table, text):
        item = table.get_children()[0]
        if table.get_children(item):  # tables with Person/Goods rows have the values on the rows below them
            item = table.get_children(item)[0]
        entry = tkinter.ttk.Entry(table.parent_frame)
        entry.insert(0, text)
        entry.editing_item_iid, entry.editing_column_index = item, 1
        before = len(messages)
        table.on_enter_pressed(SimpleNamespace(widget=entry))
        return {'shown': table.item(item, 'values')[1], 'message': messages[before][2] if len(messages) > before else None}

    gui.show_production_frame()
    out = {}
    for name, frame, amounts_table, df in [('electricity', gui.electricity_frame, gui.table_amountsEl, gui.df_Power_SC),
                                           ('district heat', gui.DH_frame, gui.table_amountsDH, gui.df_Heat_SC)]:
        out[name + ' amounts'] = {'shown': [amounts_table.item(i, 'values')[0] for i in amounts_table.get_children()],
                                  'model': [str(round(a, 2)) for a in df['Amount'][:-1]],
                                  'rows in shares table': len(last_table(frame).get_children())}
    gui.show_consumption_frame()
    shares, utilisation, prices = last_table(gui.electricity_frame), last_table(gui.transportDetails_frame), last_table(gui.economy_SC_frame)
    original_share = shares.item(shares.get_children()[0], 'values')[1]
    out['original share'] = original_share
    for name, table, text in [('share -5', shares, '-5'), ('share 120', shares, '120'), ('share abc', shares, 'abc'),
                              ('share 23,5', shares, '23,5'), ('share 7.25%', shares, '7.25%'),
                              ('utilisation 0', utilisation, '0'), ('price -3', prices, '-3'), ('price 12', prices, '12')]:
        out[name] = enter(table, text)

    co2_entry = [w for w in gui.economy_SC_frame.winfo_children() if isinstance(w, tkinter.ttk.Entry)][-1]
    co2_entry.delete(0, 'end'); co2_entry.insert(0, '-10')
    before = len(messages)
    [w for w in gui.economy_SC_frame.winfo_children() if isinstance(w, gui.FlatButton) and w.cget('text') == 'Save price changes'][-1].invoke()
    out['co2 price -10'] = {'message': messages[before][2] if len(messages) > before else None, 'co2 price now': gui.co2_price_SC}

    data_file = os.path.join(program_dir, 'data', 'Scenario Data.xlsx')
    file_before = open(data_file, 'rb').read()
    gui.df_beta_SC.loc['Heatpumps', 'Residential'] += 0.05
    gui.df_TransportFuel_SC.loc['Car', 'Electricity'] += 0.02
    before = len(messages)
    gui.update_scenario()
    out['update with wrong sums'] = {'messages': messages[before:], 'file unchanged': open(data_file, 'rb').read() == file_before}
    gui.transportFuels_sums()
    sums = gui.table_transportFuels_sum
    out['transport fuel sums'] = [[sums.item(i, 'values')[0], list(sums.item(i, 'tags'))] for i in sums.get_children()]

    gui.df_Power_SC.loc['Wind, offshore', 'Shares'] += 0.10
    gui.electricity_sums()
    out['electricity labels'] = [w.cget('text') for w in gui.electricity_frame.winfo_children() if isinstance(w, tkinter.LabelFrame)]
    return out


def scenario_files():
    """'Save scenario as', 'Reset to default scenario' and 'Open scenario', through the functions behind
    the buttons (with a file name instead of the file dialog)."""
    import openpyxl
    import pandas as pd
    import shutil
    folder = tempfile.mkdtemp()
    out = {}
    # change the scenario, then save it under a new name
    gui.df_Power_SC.loc['Coal incl. CO2-storage', 'Shares'] -= 0.10
    gui.df_Power_SC.loc['Wind, offshore', 'Shares'] += 0.10
    saved = os.path.join(folder, 'my scenario.xlsx')
    gui.save_scenario_as(saved)
    out['saved shares'] = pd.read_excel(saved, sheet_name='Electricity Production').set_index('Tech Type').loc[['Coal incl. CO2-storage', 'Wind, offshore'], 'Shares'].tolist()
    # back to the default scenario
    gui.reset_scenario()
    out['after reset'] = helpers.collect_outputs(gui)
    # open a scenario file made elsewhere: year 2035 with electrification
    other = os.path.join(folder, 'other scenario.xlsx')
    shutil.copyfile(os.path.join(program_dir, 'data', 'Default Scenario Data.xlsx'), other)
    wb = openpyxl.load_workbook(other)
    helpers.year_2035(wb)
    helpers.heat_and_transport_electrification(wb)
    wb.save(other)
    gui.open_scenario(other)
    out['after open'] = helpers.collect_outputs(gui)
    out['year shown'] = gui.entry_year.get()
    out['fresh start after open'] = helpers.run_model(program_dir)
    # a file that is not a scenario must be refused without changing anything
    bad = os.path.join(folder, 'not a scenario.xlsx')
    openpyxl.Workbook().save(bad)
    before = len(messages)
    gui.open_scenario(bad)
    out['bad file messages'] = messages[before:]
    out['after bad file'] = helpers.collect_outputs(gui)
    return out


def time_series_on_demand():
    """The time series is calculated when its page is opened, and again after the scenario changes."""
    import tkinter as tk
    out = {'calculated at startup': gui.df_hour is not None,
           'start message still shown': any('Starting pySTREAM' in str(w.cget('text')) for w in gui.root.winfo_children() if isinstance(w, tk.Label))}
    gui.show_timeSeries_frame()
    out['calculated when page opened'] = gui.df_hour is not None
    gui.update_scenario()
    out['kept after update'] = gui.df_hour is not None
    return out


def small_screen():
    """Show the window at a typical laptop size (1280 x 720) and scroll with the scrollbar and the mouse wheel."""
    root, canvas = gui.root, gui.scroll_canvas
    root.deiconify()
    root.state('normal')
    root.geometry('1280x720+0+0')
    root.update()
    out = {'horizontal scrollbar shown': gui.scroll_x.winfo_ismapped() == 1,
           'vertical scrollbar shown': gui.scroll_y.winfo_ismapped() == 1,
           'visible part at start': list(canvas.yview())}
    canvas.event_generate('<MouseWheel>', delta=-120)  # one notch down
    root.update()
    out['visible part after mouse wheel'] = list(canvas.yview())
    canvas.yview_moveto(1.0)
    root.update()
    out['visible part at the bottom'] = list(canvas.yview())
    root.withdraw()
    return out


def text_lines(text, text_font, wraplength):
    """The lines Tk shows the text in: it breaks the text at spaces so no line is longer than wraplength."""
    lines = []
    for paragraph in str(text).split('\n'):
        line = ''
        for word in paragraph.split(' '):
            longer = (line + ' ' + word) if line else word
            if line and wraplength > 0 and text_font.measure(longer) > wraplength:
                lines.append(line)
                line = word
            else:
                line = longer
        lines.append(line)
    return lines


def cut_off_text(page):
    """Texts on the shown page that do not fit in their label, button or table column."""
    import tkinter.font
    style = tkinter.ttk.Style()
    out = []
    widgets = [gui.window]
    while widgets:
        widget = widgets.pop()
        widgets.extend(widget.winfo_children())
        if not widget.winfo_ismapped():
            continue
        if isinstance(widget, (tkinter.Label, tkinter.Button, tkinter.LabelFrame)) and widget.cget('text'):
            # Only the letters themselves must fit: Tk may cut off the empty space it keeps around the text
            text_font = tkinter.font.Font(font=widget.cget('font'))
            wraplength = 0 if isinstance(widget, tkinter.LabelFrame) else widget.winfo_fpixels(widget.cget('wraplength'))
            lines = text_lines(widget.cget('text'), text_font, wraplength)
            text_size = (max(text_font.measure(line) for line in lines), text_font.metrics('linespace') * len(lines))
            room = (widget.winfo_width() - 2 * int(widget.cget('bd')), widget.winfo_height() - 2 * int(widget.cget('bd')))
            if text_size[0] > room[0] or text_size[1] > room[1]:
                out.append(f'{page}: {widget.cget("text")!r} needs {text_size[0]} x {text_size[1]} pixels, '
                           f'has {room[0]} x {room[1]}')
        elif isinstance(widget, tkinter.ttk.Treeview):
            cell_font = tkinter.font.Font(font=style.lookup(widget.cget('style') or 'Treeview', 'font'))
            heading_font = tkinter.font.Font(font=style.lookup('Treeview.Heading', 'font'))
            padding = 2 * gui.px(4)
            columns = list(widget['columns'])
            for i, column in enumerate(columns):
                column_width = widget.column(column, 'width')
                if 'headings' in str(widget.cget('show')) and heading_font.measure(widget.heading(column, 'text')) + padding > column_width:
                    out.append(f'{page}: heading {widget.heading(column, "text")!r} does not fit in {column_width} pixels')
                rows = list(widget.get_children())
                rows += [child for row in rows for child in widget.get_children(row)]
                for row in rows:
                    values = widget.item(row, 'values')
                    if i < len(values) and cell_font.measure(str(values[i])) + padding > column_width:
                        out.append(f'{page}: {values[i]!r} in column {column!r} does not fit in {column_width} pixels')
    return out


def screen_layout():
    """Start pySTREAM on a screen of another size (see SIMULATED_SCREENS) and look at every page."""
    root = gui.root
    root.deiconify()
    root.update()
    out = {'scale': gui.SCALE, 'layout': [gui.ww, gui.wh], 'screen': list(SIMULATED_SCREENS[check]),
           'tk scaling': float(root.tk.call('tk', 'scaling')), 'table row height': gui.style.lookup('Treeview', 'rowheight'),
           'cut off text': []}
    pages = {'graphs': gui.graphs_button_action, 'consumption': gui.consumption_button_action,
             'production': gui.production_button_action, 'base year': gui.by_button_action,
             'scenario': gui.scenario_button_action, 'transport': gui.transport_button_action,
             'CO2': gui.CO2_button_action, 'economy': gui.economy_button_action}
    for page, show in pages.items():
        show()
        root.update()
        out['cut off text'] += cut_off_text(page)
    out['horizontal scrollbar shown'] = gui.scroll_x.winfo_ismapped() == 1
    out['vertical scrollbar shown'] = gui.scroll_y.winfo_ismapped() == 1
    root.withdraw()
    return out


def button_click():
    """Click the 'CO2 emissions' menu button with the mouse: once normally, once letting go outside the button."""
    root, button = gui.root, gui.CO2_button
    root.deiconify()
    gui.graphs_button_action()
    root.update()
    out = {'colours': [button.cget('bg'), button.cget('fg')]}
    button.event_generate('<ButtonPress-1>', x=5, y=5)
    root.update()
    out['state while pressed'] = str(button.cget('state'))
    out['colours while pressed'] = [button.cget('activebackground'), button.cget('activeforeground')]
    button.event_generate('<ButtonRelease-1>', x=-20, y=5)  # let go outside the button
    root.update()
    out['CO2 page after letting go outside'] = gui.CO2_frame.winfo_ismapped() == 1
    button.event_generate('<ButtonPress-1>', x=5, y=5)
    button.event_generate('<ButtonRelease-1>', x=5, y=5)
    root.update()
    out['CO2 page after click'] = gui.CO2_frame.winfo_ismapped() == 1
    out['state after click'] = str(button.cget('state'))
    root.withdraw()
    return out


CHECKS = {
    'button_click': button_click,
    'screen_laptop': screen_layout,
    'screen_design': screen_layout,
    'screen_large': screen_layout,
    'screen_tiny': screen_layout,
    'small_screen': small_screen,
    'time_series_on_demand': time_series_on_demand,
    'scenario_files': scenario_files,
    'input_validation': input_validation,
    'file_open_in_excel': file_open_in_excel,
    'plots': plots,
    'economy_page': economy_page,
    'sankey_image_fails': sankey_buttons,
    'sankey_image_hangs': sankey_buttons,
    'close_window': close_window,
    'chp_surplus_warning': chp_surplus_warning,
    'start_from_other_folder': start_from_other_folder,
    'co2_totals': co2_totals,
    'update_matches_fresh_start': update_matches_fresh_start,
    'update_time_series_button': update_time_series_button,
    'update_scenario_year': update_scenario_year,
}

result = CHECKS[check]()
result['messages'] = messages
with open(output_file, 'w', encoding='utf-8') as f:
    json.dump(result, f, default=helpers.number)

# Exit right away: the threads Kaleido starts to render the Sankey images sometimes
# keep Python from shutting down normally, which made these checks hang at the end.
os._exit(0)
