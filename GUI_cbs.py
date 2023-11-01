import tkinter.filedialog as filedialog
import tkinter as tk
import subprocess
import openpyxl
import lxml  
from excel_verwerking import excel

master = tk.Tk()
master.title("CBS IHG rapportage")

def input1():
    print("ik ben in functie input1")
    input_path = filedialog.askopenfilename()
    print("input_path:", input_path)
    input_entry1.delete(1, tk.END)  # Remove current text in entry
    input_entry1.insert(0, input_path)  # Insert the 'path'

def input2():
    print("ik ben in functie input2")
    input_path = filedialog.askopenfilename()
    print("input_path:", input_path)
    input_entry2.delete(0, tk.END)
    input_entry2.insert(0, input_path)

def input3():
    print("ik ben in functie input3")
    input_path = filedialog.askopenfilename()
    print("input_path:", input_path)
    input_entry3.delete(0, tk.END)
    input_entry3.insert(0, input_path)

def input4():
    print("ik ben in functie input4")
    input_path = filedialog.askopenfilename()
    print("input_path:", input_path)
    input_entry4.delete(0, tk.END)
    input_entry4.insert(0, input_path)

def input5():
    print("ik ben in functie input5")
    input_path = filedialog.askopenfilename()
    print("input_path:", input_path)
    input_entry5.delete(0, tk.END)
    input_entry5.insert(0, input_path)

def input6():
    print("ik ben in functie input6")
    input_path = filedialog.askopenfilename()
    print("input_path:", input_path)
    input_entry6.delete(0, tk.END)
    input_entry6.insert(0, input_path)

def output():
    print("ik ben in functie output")
    output_path = filedialog.asksaveasfilename(defaultextension=".xlsx")
    print("output_path:", output_path)
    output_entry.delete(1, tk.END)  # Remove current text in entry
    output_entry.insert(0, output_path)  # Insert the 'path'

def press_button():
    print("ik ben in functie press_button")
    input_paths = [input_entry1.get(), input_entry2.get(), input_entry3.get(), input_entry4.get(), input_entry5.get(), input_entry6.get()]
    output_path = output_entry.get()
    try:
        subprocess.run(["python", "begin.py", *input_paths, output_path], check=True)
        print("ik voer panda-script uit")
    except subprocess.CalledProcessError as e:
        print(f"fout bij het uitvoeren van het panda-script: {e}")
        print("ik kan geen panda-script uitvoeren")
    excel(input_paths, output_path)

top_frame = tk.Frame(master)
bottom_frame = tk.Frame(master)
line = tk.Frame(master, height=1, width=400, bg="grey80", relief='groove')

input_label1 = tk.Label(top_frame, text="Exact ICV:")
input_entry1 = tk.Entry(top_frame, width=40)
browse1 = tk.Button(top_frame, text="Browse", command=input1)

input_label2 = tk.Label(top_frame, text="Exact overzicht boekingen:")
input_entry2 = tk.Entry(top_frame, width=40)
browse2 = tk.Button(top_frame, text="Browse", command=input2)

input_label3 = tk.Label(top_frame, text="Exact overzicht relaties:")
input_entry3 = tk.Entry(top_frame, width=40)
browse3 = tk.Button(top_frame, text="Browse", command=input3)

input_label4 = tk.Label(top_frame, text="PIM producten gewicht export:")
input_entry4 = tk.Entry(top_frame, width=40)
browse4 = tk.Button(top_frame, text="Browse", command=input4)

input_label5 = tk.Label(top_frame, text="vendit overzicht ingekochte producten:")
input_entry5 = tk.Entry(top_frame, width=40)
browse5 = tk.Button(top_frame, text="Browse", command=input5)

input_label6 = tk.Label(top_frame, text="vendit producten export:")
input_entry6 = tk.Entry(top_frame, width=40)
browse6 = tk.Button(top_frame, text="Browse", command=input6)

output_label = tk.Label(bottom_frame, text="Output File Path:")
output_entry = tk.Entry(bottom_frame, width=40)
browse_output = tk.Button(bottom_frame, text="Browse", command=output)

begin_button = tk.Button(bottom_frame, text='Begin!', command=press_button)

top_frame.pack(side=tk.TOP)
line.pack(pady=10) 
bottom_frame.pack(side=tk.BOTTOM)

input_label1.grid(row=0, column=0, padx=10, pady=5, sticky='w')
input_entry1.grid(row=0, column=1, padx=10, pady=5, columnspan=2)
browse1.grid(row=0, column=3, padx=10, pady=5)

input_label2.grid(row=1, column=0, padx=10, pady=5, sticky='w')
input_entry2.grid(row=1, column=1, padx=10, pady=5, columnspan=2)
browse2.grid(row=1, column=3, padx=10, pady=5)

input_label3.grid(row=2, column=0, padx=10, pady=5, sticky='w')
input_entry3.grid(row=2, column=1, padx=10, pady=5, columnspan=2)
browse3.grid(row=2, column=3, padx=10, pady=5)

input_label4.grid(row=3, column=0, padx=10, pady=5, sticky='w')
input_entry4.grid(row=3, column=1, padx=10, pady=5, columnspan=2)
browse4.grid(row=3, column=3, padx=10, pady=5)

input_label5.grid(row=4, column=0, padx=10, pady=5, sticky='w')
input_entry5.grid(row=4, column=1, padx=10, pady=5, columnspan=2)
browse5.grid(row=4, column=3, padx=10, pady=5)

input_label6.grid(row=5, column=0, padx=10, pady=5, sticky='w')
input_entry6.grid(row=5, column=1, padx=10, pady=5, columnspan=2)
browse6.grid(row=5, column=3, padx=10, pady=5)

output_label.grid(row=6, column=0, padx=10, pady=5, sticky='w')
output_entry.grid(row=6, column=1, padx=10, pady=5, columnspan=2)
browse_output.grid(row=6, column=3, padx=10, pady=5)

begin_button.grid(row=7, column=0, columnspan=4, pady=20)

master.mainloop()