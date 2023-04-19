### Import the required libraries
from tkinter import *
from functions import *

### set frames and base grid
# create instance
root = Tk()
# set base dimensions and features
root.geometry("470x500")
root.title('Bill_Bot v1')

### set dimensions for frames
task_frame = LabelFrame(root, text="Tasks", padx=2, pady=2)
task_frame.grid(rowspan=8, row=0, column=0)

menu_frame = LabelFrame(root, padx=5, pady=5)
menu_frame.grid(rowspan=1, row=1, column=1)

### title
label = Label(root, text="Welcome!")
label.grid(row=0, column=1)

### Create dropdown Menus
listbox_daily = Listbox(task_frame, width=40, height=24, selectmode=MULTIPLE)

# Inserting the listbox items
listbox_daily.insert(1, "d-Posted/Unposted")
listbox_daily.insert(2, "d-Focus File")
listbox_daily.insert(3, "d-Overdue Invoices")
listbox_daily.insert(4, "d-All Daily")

listbox_daily.insert(5, "w-Contract Mapping")
listbox_daily.insert(6, "w-AR Report")
listbox_daily.insert(7, "w-PIE Import")
listbox_daily.insert(8, "w-Budget V Spend")
listbox_daily.insert(9, "w-SST to SBLIW")

listbox_daily.insert(10, "m-Quick Pulse")
listbox_daily.insert(11, "m-MI45 Reminder")
listbox_daily.insert(12, "m-FIWLR")
listbox_daily.insert(13, "m-FIWLR w/Miscodes")
listbox_daily.insert(14, "m-CP Actuals")
listbox_daily.insert(15, "m-Cadence Files")
listbox_daily.insert(16, "m-EOM AR Report")
listbox_daily.insert(17, "m-OEM Files")
listbox_daily.insert(18, "m-Odd Day")
listbox_daily.insert(19, "m-Billing Audit")
listbox_daily.insert(20, "m-Cummulative Overdue")
listbox_daily.insert(21, "m-Planner Checks")

listbox_daily.pack()

### Function to process selected requests
tasks_list = []
def submit_requests():
    label = Label(root, text="")
    for i in listbox_daily.curselection():
        tasks_list.append(listbox_daily.get(i))
    if tasks_list == []:
        label = Label(root, text="Please select your tasks and try again...")
        label.grid(row=0, column=1)
    else:
        select_docs()
        label = Label(root, text="Your requests are being processed...")
        label.grid(row=0, column=1)
        perform_tasks(task_list=tasks_list)

# button to process reports
submit_button = Button(\
    menu_frame,\
    text="Prepare Reports",\
    padx=50,\
    command=submit_requests)\
    .grid(row=4,column=10)

### needed for app
root.mainloop()
