# Python program to create
# a file explorer in Tkinter

# import all components
# from the tkinter library
from tkinter import *
from urllib.request import urlopen
import datetime
# import filedialog module
from tkinter import filedialog, messagebox
from send_whatsapp_message import SendMessage

# Function for opening the
# file explorer window

def license_check():
    try:
        res = urlopen('http://just-the-time.appspot.com/')
        result = res.read().strip()
        result_str = result.decode('utf-8')
        if datetime.datetime.strptime(result_str, '%Y-%m-%d %H:%M:%S') > datetime.datetime.strptime('2020-07-20 23:59:00',
                                                                                                    '%Y-%m-%d %H:%M:%S'):
            messagebox.showerror("Error", "You cannot execute the app, check the license")
            exit()
    except:
        messagebox.showerror("Error", "Check the connection")
        exit()

def browseDataFiles():
    filename = filedialog.askopenfilename(title="Select a File",
                                          filetypes=(("Excel files",
                                                      "*.xlsx*"),
                                                     ("all files",
                                                      "*.*")))
    # Change label contents
    label_data_explorer.configure(text= filename)
    print(label_data_explorer["text"])

def browseMessageFiles():
    filename = filedialog.askopenfilename(title="Select a File",
                                          filetypes=(("Text files",
                                                      "*.txt*"),
                                                     ("all files",
                                                      "*.*")))

    # Change label contents
    label_message_explorer.configure(text= filename)
# Create the root window

def start_send():
    mobile_no = mobile_no_txt.get()
    send_message = SendMessage(mobile_no, label_data_explorer["text"], label_message_explorer["text"])
    send_message.send_massage()


# license_check()
window = Tk()

# Set window title
window.title('Send WhatsApp Messages')

# Set window size
window.geometry("800x300")

# Set window background color
window.config()


# Create a File Explorer label
label_data_explorer = Label(window,
                            width=100, height=2,
                            fg="blue")
label_message_explorer = Label(window,
                            width=100, height=2,
                            fg="blue")

button_data_explore = Button(window,
                        text="  Select Data File ",
                        command=browseDataFiles)

button_message_explore = Button(window,
                        text="Select Message File",
                        command=browseMessageFiles)

button_start = Button(window,
                     text="Start",
                     command=start_send)

# Grid method is chosen for placing
# the widgets at respective positions
# in a table like structure by
# specifying rows and columns

# mobile_no = "+86 153 7240 6855"
label_mobile_no = Label(window,
                            width= 15,
                            height=2,
                            fg="blue", text="Mobile No")
label_mobile_no.grid(column=0, row=0)
mobile_no = StringVar()
mobile_no_txt = Entry(window, width = 30, textvariable = mobile_no)
mobile_no_txt.grid(column = 1, row = 0)

button_data_explore.grid(column=0, row=2)
label_data_explorer.grid(column=1, row=2)

button_message_explore.grid(column=0, row=3)
label_message_explorer.grid(column=1, row=3)

button_start.grid(column=1, row=4)




# Let the window wait for any events
window.mainloop()