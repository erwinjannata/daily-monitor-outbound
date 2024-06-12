import tkinter as tk
from tkinter import filedialog, ttk
from tkinter.messagebox import showinfo
from tkcalendar import DateEntry
import threading
import os
from functions.function import grouping, daily_monitor

root = tk.Tk()
root.configure(bg="white")
root.geometry("350x475")
root.title("JNE AMI Outbound App")
root.resizable(0, 0)

file_data = ""
file_report = ""

file_data_name = tk.StringVar()
file_report_name = tk.StringVar()
mode = tk.StringVar()
over_month = tk.IntVar()


def load_data():
    global file_data
    file_data = filedialog.askopenfilename(filetypes=(
        ("Excel Workbook", "*.xlsx"),
        ("All Files", "*.*"),
    ))
    file_data_name.set(os.path.split(file_data)[1])


def load_master_report():
    global file_report
    file_report = filedialog.askopenfilename(filetypes=(
        ("Excel Workbook", "*.xlsx"),
        ("All Files", "*.*"),
    ))
    file_report_name.set(os.path.split(file_report)[1])


def combine_process():
    mode = combo_box.current()
    saved_as = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[
                                            ("Excel Workbook (.xlsx)", "*.xlsx")])
    global date
    date = calendar.get()

    if os.path.exists(file_data) and saved_as:
        progressbar.start()
        if mode == 0 and os.path.exists(file_report):
            daily_monitor(file_data=file_data, file_report=file_report,
                          date=date, saved_as=saved_as, over_month=over_month.get())
        elif mode == 0 and not os.path.exists(file_report):
            showinfo(title="Message",
                     message="File report terbaru tidak ditemukan")
        elif mode == 1:
            grouping(file_data=file_data, date=date,
                     save_grouping=True, saved_as=saved_as)
        else:
            showinfo(title="Message",
                     message="Input tidak valid")
    elif not file_data:
        showinfo(title="Message",
                 message="File data kosong")
    elif not saved_as:
        showinfo(title="Message",
                 message="Pilih lokasi penyimpanan file yang valid")
    else:
        showinfo(title="Message",
                 message="Cek kembali excel yang dipilih!")
    combo_box.state(['!disabled'])
    calendar.state(['!disabled'])
    btn1.state(['!disabled'])
    btn2.state(['!disabled'])
    check_label.config(state="normal")
    over_month.set(0)
    combine_btn.state(['!disabled'])


def start_combine_thread(event):
    global combine_thread
    combine_thread = threading.Thread(target=combine_process)
    combine_thread.daemon = True
    combo_box.state(['disabled'])
    calendar.state(['disabled'])
    btn1.state(['disabled'])
    btn2.state(['disabled'])
    check_label.config(state="disabled")
    combine_btn.state(['disabled'])
    combine_thread.start()
    root.after(20, check_combine_thread)


def check_combine_thread():
    if combine_thread.is_alive():
        root.after(20, check_combine_thread)
    else:
        progressbar.stop()


# GUI
combo_label = ttk.Label(root, text="1. Proses",
                        background="white", font="calibri 11 bold").pack(fill="x", padx=10, pady=5)
combo_box = ttk.Combobox(root, textvariable=mode)
combo_box['value'] = ('Daily Monitoring Outbound',
                      'Grouping Data Daily Monitor')
combo_box.pack(pady=10, padx=10, fill='both')
combo_box.current(0)

label1 = ttk.Label(root, text="2. Tanggal Data (H+0)", background="white", font="calibri 11 bold").pack(
    fill="x", padx=10, pady=5)

calendar = DateEntry(root, selectmode='day', locale='en_US',
                     date_pattern='M/d/yyyy', weekendbackground='white', weekendforeground='black')
calendar.pack(pady=10, padx=10, fill='both')

check_label = tk.Checkbutton(
    root, text="Report bulan sebelumnya", background="white", variable=over_month, onvalue=1, offvalue=0)
check_label.pack(pady=5, padx=10, anchor="w")

label2 = ttk.Label(root, text="3. File Data", background="white", font="calibri 11 bold").pack(
    fill="x", padx=10, pady=5)

label_name1 = ttk.Label(root, textvariable=file_data_name, background="white").pack(
    fill="x", padx=10, pady=5)


btn1 = ttk.Button(root, text="Pilih File", command=load_data, state=tk.NORMAL)
btn1.pack(fill="x", padx=10, pady=5)

label3 = ttk.Label(root, text="4. File Daily Monitor Terbaru", background="white", font="calibri 11 bold").pack(
    fill="x", padx=10, pady=5)

label_name2 = ttk.Label(root, textvariable=file_report_name, background="white").pack(
    fill="x", padx=10, pady=5)

btn2 = ttk.Button(root, text="Pilih File",
                  command=load_master_report, state=tk.NORMAL)
btn2.pack(fill="x", padx=10, pady=5)

separator = ttk.Separator(root, orient='horizontal').pack(
    fill='x', pady=5, padx=10)

combine_btn = ttk.Button(root, text="Proses",
                         command=lambda: start_combine_thread(None), state=tk.NORMAL)
combine_btn.pack(fill="x", padx=10, pady=10)

progressbar = ttk.Progressbar(root, mode='indeterminate', orient='horizontal')
progressbar.pack(fill='x', padx=10, pady=10)

root.mainloop()
