import tkinter as tk
import subprocess

def launch_app(app_path):
    subprocess.Popen([app_path], shell=True)

root = tk.Tk()
root.title("Rocs one time Task")

# Function to launch different apps
def launch_app1():
    launch_app("onetimetask.exe")

def launch_app2():
    launch_app("RMS.exe")

def launch_app3():
    launch_app("EDI_SMS_sep_1.exe")

def launch_app4():
    launch_app("EDI_SMS_step_2.exe")

def launch_app5():
    launch_app("Invoice_Code_Usage_Translaion.exe")

# Create buttons
button1 = tk.Button(root, text="Invoice Code Usage SMS", command=launch_app1, font=("Helvetica", 14))
button2 = tk.Button(root, text="Invoice Code Usage RMS", command=launch_app2, font=("Helvetica", 14))
button3 = tk.Button(root, text="EDI SMS Step 1", command=launch_app3, font=("Helvetica", 14))
button4 = tk.Button(root, text="EDI SMS Step 2", command=launch_app4, font=("Helvetica", 14))
button5 = tk.Button(root, text="Invoice Code Usage\nTranslation", command=launch_app5, font=("Helvetica", 14))

# Configure button sizes
button_width = 20
button_height = 2

button1.config(width=button_width, height=button_height)
button2.config(width=button_width, height=button_height)
button3.config(width=button_width, height=button_height)
button4.config(width=button_width, height=button_height)
button5.config(width=button_width, height=button_height)

# Grid layout for buttons
button1.grid(row=0, column=0, padx=10, pady=10)
button2.grid(row=0, column=1, padx=10, pady=10)
button3.grid(row=1, column=0, padx=10, pady=10)
button4.grid(row=1, column=1, padx=10, pady=10)
button5.grid(row=2, column=0, columnspan=2, padx=10, pady=10)

root.mainloop()
