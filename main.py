import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
from tkinter import ttk
from PIL import ImageTk, Image
from tkcalendar import DateEntry


form = tk.Tk()
form.title("Excel tools Form")
form.geometry("400x280")

tab_parent = ttk.Notebook(form)

tab1 = ttk.Frame(tab_parent)
tab2 = ttk.Frame(tab_parent)
tab3 = ttk.Frame(tab_parent)
tab_parent.add(tab1, text="Details Tool Usage")
tab_parent.add(tab2, text="Summary")
tab_parent.add(tab3, text="Access Rights")

excel_tools = tk.StringVar()
excel_tools.set("Tool1 Tool2 Tool3 Tool4 Tool5 Tool6 Tool7")

users = tk.StringVar()
users.set("User1 User2 User3")

def generate1():
    info = ""
    info += excel_tools_entry_tab1.get() + "\n"
    info += date_from_entry_tab1.get() + "\n"
    info += date_to_entry_tab1.get()
    tk.messagebox.showinfo("Info", info)
    
def generate2():
    info = ""
    info += excel_tools_entry_tab2.get() + "\n"
    info += date_from_entry_tab2.get() + "\n"
    info += date_to_entry_tab2.get()
    tk.messagebox.showinfo("Info", info)
    
def generate3():
    info = ""
    info += excel_tools_entry_tab3.get() + "\n"
    info += users_entry_tab3.get()
    tk.messagebox.showinfo("Info", info)


def excel_tools_window(entry, _list):
    def add():
        selections = lstbox.curselection()
        entry.config(state="normal")
        entry.delete(0, tk.END)
        new_list = ";".join([lstbox.get(i) for i in selections])
        entry.insert(0, new_list)
        entry.config(state="readonly")
    window = tk.Toplevel(form)
    lstbox = tk.Listbox(window, listvariable=_list, selectmode="multiple", width=20, height=10)
    lstbox.grid(column=0, row=0, columnspan=2)
    btn = ttk.Button(window, text="Add", command=add)
    btn.grid(column=1, row=1)

# === WIDGETS FOR TAB ONE
excel_tools_label_tab1 = tk.Label(tab1, text="Excel tools:")
date_from_label_tab1 = tk.Label(tab1, text="Date from:")
date_to_label_tab1 = tk.Label(tab1, text="Date to:")

excel_tools_entry_tab1 = tk.Entry(tab1, state="readonly")
date_from_entry_tab1 = DateEntry(tab1, width=12, background='darkblue',
                    foreground='white', borderwidth=2)
date_to_entry_tab1 = DateEntry(tab1, width=12, background='darkblue',
                    foreground='white', borderwidth=2)

imgLabelTabOne = tk.Label(tab1)

generate_button_tab1 = tk.Button(tab1, text="Generate", command=generate1)
select_button_tab1 = tk.Button(tab1, text="Select", command=lambda: excel_tools_window(excel_tools_entry_tab1, excel_tools))

# === ADD WIDGETS TO GRID ON TAB ONE
excel_tools_label_tab1.grid(row=0, column=0, padx=15, pady=15)
excel_tools_entry_tab1.grid(row=0, column=1, padx=15, pady=15)

date_from_label_tab1.grid(row=1, column=0, padx=15, pady=15)
date_from_entry_tab1.grid(row=1, column=1, padx=15, pady=15)

date_to_label_tab1.grid(row=2, column=0, padx=15, pady=15)
date_to_entry_tab1.grid(row=2, column=1, padx=15, pady=15)

imgLabelTabOne.grid(row=0, column=2, rowspan=3, padx=15, pady=15)

generate_button_tab1.grid(row=3, column=0, padx=15, pady=15)
select_button_tab1.grid(row=0, column=2, padx=15, pady=15)

# === WIDGETS FOR TAB TWO
excel_tools_label_tab2 = tk.Label(tab2, text="Excel tools:")
date_from_label_tab2 = tk.Label(tab2, text="Date from:")
date_to_label_tab2 = tk.Label(tab2, text="Date to:")

excel_tools_entry_tab2 = tk.Entry(tab2)
date_from_entry_tab2 = DateEntry(tab2, width=12, background='darkblue',
                    foreground='white', borderwidth=2)
date_to_entry_tab2 = DateEntry(tab2, width=12, background='darkblue',
                    foreground='white', borderwidth=2)

imgLabelTabOne = tk.Label(tab2)

generate_button_tab2 = tk.Button(tab2, text="Generate", command=generate2)
select_excel_tools_button_tab2 = tk.Button(tab2, text="Select", command=lambda: excel_tools_window(excel_tools_entry_tab2, excel_tools))

# === ADD WIDGETS TO GRID ON TAB TWO
excel_tools_label_tab2.grid(row=0, column=0, padx=15, pady=15)
excel_tools_entry_tab2.grid(row=0, column=1, padx=15, pady=15)

date_from_label_tab2.grid(row=1, column=0, padx=15, pady=15)
date_from_entry_tab2.grid(row=1, column=1, padx=15, pady=15)

date_to_label_tab2.grid(row=2, column=0, padx=15, pady=15)
date_to_entry_tab2.grid(row=2, column=1, padx=15, pady=15)

imgLabelTabOne.grid(row=0, column=2, rowspan=3, padx=15, pady=15)

generate_button_tab2.grid(row=3, column=0, padx=15, pady=15)
select_excel_tools_button_tab2.grid(row=0, column=2, padx=15, pady=15)

# === WIDGETS FOR TAB TWO
excel_tools_label_tab3 = tk.Label(tab3, text="Excel tools:")
users_label_tab3 = tk.Label(tab3, text="Users:")

excel_tools_entry_tab3 = tk.Entry(tab3)
users_entry_tab3 = tk.Entry(tab3)

imgLabelTabOne = tk.Label(tab3)

generate_button_tab3 = tk.Button(tab3, text="Generate", command=generate3)
select_excel_tools_button_tab3 = tk.Button(tab3, text="Select", command=lambda: excel_tools_window(excel_tools_entry_tab3, excel_tools))
select_users_button_tab3 = tk.Button(tab3, text="Select", command=lambda: excel_tools_window(users_entry_tab3, users))

# === ADD WIDGETS TO GRID ON TAB TWO
excel_tools_label_tab3.grid(row=0, column=0, padx=15, pady=15)
excel_tools_entry_tab3.grid(row=0, column=1, padx=15, pady=15)

users_label_tab3.grid(row=1, column=0, padx=15, pady=15)
users_entry_tab3.grid(row=1, column=1, padx=15, pady=15)

imgLabelTabOne.grid(row=0, column=2, rowspan=3, padx=15, pady=15)

generate_button_tab3.grid(row=2, column=0, padx=15, pady=15)
select_excel_tools_button_tab3.grid(row=0, column=2, padx=15, pady=15)
select_users_button_tab3.grid(row=1, column=2, padx=15, pady=15)

tab_parent.pack(expand=1, fill='both')

form.mainloop()