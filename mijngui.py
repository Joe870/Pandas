import tkinter.filedialog as filedialog
import tkinter as tk



# functie voor de input
def select_file(entry_widget):
    path = filedialog.askopenfilename()
    entry_widget.delete(0, tk.END)
    entry_widget.insert(0, path)
    
def select_output_file(entry_widget):
    path = filedialog.asksaveasfilename(defaultextension=".xslx", filetypes=[("excel-bestanden", "*.xslx")])
    entry_widget.delete(0, tk.END)
    entry_widget.insert(0, path) 

# begin
master = tk.Tk()
master.title("CBS IHG rapportage")
top_frame = tk.Frame(master)
bottom_frame = tk.Frame(master)
line = tk.Frame(master, height=1, width=400, bg="grey80", relief='groove')

browse_input1 = tk.Button(top_frame, text="Browse", command=lambda: select_file(input_entry1))
browse_input2 = tk.Button(top_frame, text="Browse", command=lambda: select_file(input_entry2))
browse_input3 = tk.Button(top_frame, text="Browse", command=lambda: select_file(input_entry3))
browse_input4 = tk.Button(top_frame, text="Browse", command=lambda: select_file(input_entry4))
browse_input5 = tk.Button(top_frame, text="Browse", command=lambda: select_file(input_entry5))
browse_input6 = tk.Button(top_frame, text="Browse", command=lambda: select_file(input_entry6))
browse_output = tk.Button(bottom_frame, text="Browse", command=lambda: select_output_file(output_entry))

input_label1 = tk.Label(top_frame, text="exact ICV:")
input_entry1 = tk.Entry(top_frame, width=40)
browse_input1 = tk.Button(top_frame, text="Browse", command=lambda: select_file(input_entry1))

input_label2 = tk.Label(top_frame, text="exact overzicht boekingen:")
input_entry2 = tk.Entry(top_frame, width=40)
browse_input2 = tk.Button(top_frame, text="Browse", command=lambda: input(input_entry2))

input_label3 = tk.Label(top_frame, text="exact overzicht relaties:")
input_entry3 = tk.Entry(top_frame, width=40)
browse_input3 = tk.Button(top_frame, text="Browse", command=lambda: input(input_entry3))

input_label4 = tk.Label(top_frame, text="vendit overzicht ingekochte producten:")
input_entry4 = tk.Entry(top_frame, width=40)
browse_input4 = tk.Button(top_frame, text="Browse", command=lambda: input(input_entry4))

input_label5 = tk.Label(top_frame, text="vendit producten export:")
input_entry5 = tk.Entry(top_frame, width=40)
browse_input5 = tk.Button(top_frame, text="Browse", command=lambda: input(input_entry5))

input_label6 = tk.Label(top_frame, text="PIM producten gewichten export:")
input_entry6 = tk.Entry(top_frame, width=40)
browse_input6 = tk.Button(top_frame, text="Browse", command=lambda: input(input_entry6))

output_label = tk.Label(bottom_frame, text="Output excel bestand:")
output_entry = tk.Entry(bottom_frame, width=40)
browse_output = tk.Button(bottom_frame, text="Browse", command=lambda: input(output_entry))

begin_button = tk.Button(bottom_frame, text='Begin!')

input_label1.grid(row=0, column=0, pady=5)
input_entry1.grid(row=0, column=1, pady=5)
browse_input1.grid(row=0, column=2, pady=5)

input_label2.grid(row=1, column=0, pady=5)
input_entry2.grid(row=1, column=1, pady=5)
browse_input2.grid(row=1, column=2, pady=5)

input_label3.grid(row=2, column=0, pady=5)
input_entry3.grid(row=2, column=1, pady=5)
browse_input3.grid(row=2, column=2, pady=5)

input_label4.grid(row=3, column=0, pady=5)
input_entry4.grid(row=3, column=1, pady=5)
browse_input4.grid(row=3, column=2, pady=5)

input_label5.grid(row=4, column=0, pady=5)
input_entry5.grid(row=4, column=1, pady=5)
browse_input5.grid(row=4, column=2, pady=5)

input_label6.grid(row=5, column=0, pady=5)
input_entry6.grid(row=5, column=1, pady=5)
browse_input6.grid(row=5, column=2, pady=5)

begin_button.grid(row=1, columnspan=3, pady=20, sticky="ew")

top_frame.pack(side=tk.TOP)
line.pack(pady=10)
bottom_frame.pack(side=tk.BOTTOM)

master.mainloop()