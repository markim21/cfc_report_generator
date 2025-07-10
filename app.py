from nicegui import ui, events, native
from nicegui.events import ValueChangeEventArguments
from counts import generate_counts_reports
from task_reports import generate_task_report

filename = ""
output = "output.xlsx"
func = "task"

def set_file(x: events.ValueChangeEventArguments):
    global filename
    filename = x.value

def set_output(x: events.ValueChangeEventArguments):
    global output
    output = x.value

def set_func(x: events.ValueChangeEventArguments):
    global func
    func = x.value 

def report_function(*args):
    if func == "Task Report": return generate_task_report(args)
    if func == "Facility Utilization Report": return generate_counts_reports(args)

def generate_wrapper():
    global filename, output
    if filename == "":
        ui.notify("List a file to process.")
    elif filename[-4:] != ".csv":
        ui.notify("This tool can only process .csv files from Connect2. Please change the input file name to include .csv")
    elif output[-5:] != ".xlsx":
        ui.notify("This tool can only output .xlsx files. Please change the output name to include .xlsx")
    else:
        try:
            success = report_function(filename, output)
        except FileNotFoundError:
            ui.notify(filename + " not found in current folder. Make sure the file is in the same folder as the executable file.")
        finally: 
            if success:
                ui.notify('Report generated.')
            else:
                ui.notify('Error generating report.')

ui.page_title('Connect2 Reports Generator')
ui.label('This tool generates a report on fitness center utilization based on Connect2 Filtered Counts Reports.')
ui.label('Use this tool as follows:')
ui.label('1. Generate and download a Filtered Counts Report from Connect2. This should be a .csv file and only include fitness centers (NOT issue rooms or courts)')
ui.label('2. Save the Connect2 file in the same file directory of this tool.')
ui.label('3. Type the name of the file in the box below. Include the file extension (".csv") in the name. For example, a file could be named "input.csv"')
ui.label('4. (Optional) Type the name of the resulting report in the box below. Include the file extension (".xlsx") in the name. The default name is "output.xlsx"')

with ui.row():
    with ui.column():
        ui.input(label='Name of input file:', on_change=lambda x: set_file(x))
        ui.input(label='Name of output file:', on_change = lambda x: set_output(x))
        ui.select(['Facility Utilization Report', 'Task Report'], value='Task Report', on_change=lambda x: set_func(x))
        ui.button("press to generate report", on_click=lambda: generate_wrapper())

ui.run(title='Connect2 Reports Generator', reload=False, native=True)

# nicegui-pack --onefile --windowed --name "testapp" app.py