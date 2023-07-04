import tkinter as tk
from test import gui_handle

def submit(email, app_key):
    print("Email:", email)
    print("App Key:", app_key)
    gui_handle(email, app_key)
    success_label.config(text="Loaded Successfully!")  # Update the label text
    root.after(4000, root.destroy)  # Schedule window closure after 4000 milliseconds (4 seconds)

root = tk.Tk()
root.title("E-Mail Extractor")

# Set the initial size of the window
window_width = 400
window_height = 300
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x = (screen_width - window_width) // 2
y = (screen_height - window_height) // 2
root.geometry(f"{window_width}x{window_height}+{x}+{y}")

# Creating the email input field and label
email_label = tk.Label(root, text="Email:")
email_label.pack(pady=5)  # Adding vertical padding
email_entry = tk.Entry(root)
email_entry.pack(pady=5)  # Adding vertical padding

# Creating the app key input field and label
app_key_label = tk.Label(root, text="App Key:")
app_key_label.pack(pady=5)  # Adding vertical padding
app_key_entry = tk.Entry(root, show="*")  # Use show="*" to hide the entered characters (e.g., for password input)
app_key_entry.pack(pady=5)  # Adding vertical padding

def submit_wrapper():
    email = email_entry.get()
    app_key = app_key_entry.get()
    submit(email, app_key)

# Creating the submit button
submit_button = tk.Button(root, text="Submit", command=submit_wrapper)
submit_button.pack(pady=10)  # Adding vertical padding

# Creating the success label
success_label = tk.Label(root, text="")
success_label.pack(pady=10)  # Adding vertical padding

root.mainloop()
