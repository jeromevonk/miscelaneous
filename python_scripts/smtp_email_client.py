"""
email_example4.py
~~~~~~
 
Creates a simple GUI for sending emails
"""

import threading
import smtplib
import tkinter
from tkinter import ttk
from email.message import EmailMessage

class EmailSender(ttk.Frame):
    """The gui and functions."""
    def __init__(self, parent, *args, **kwargs):
        ttk.Frame.__init__(self, parent, *args, **kwargs)
        self.smtpObj = 0
        self.root = parent
        self.init_gui()
 
    def on_quit(self):
        """Exits program."""
        quit()
 
    def connect(self):
        """Connects to smtp server"""
        def connect_to_smtp():
            # Create a SMPT object and call 'hello'
            _smtp    = self.smtp_entry.get()
            self.smtpObj = smtplib.SMTP(_smtp, 587)
            self.smtpObj.ehlo() 
            
            # Start TLS
            self.smtpObj.starttls()
            
            # Login
            _from    = self.from_entry.get()
            _pswd    = self.pswd_entry.get()
            
            self.smtpObj.login(_from, _pswd)
            
            self.answer_label['text'] = 'Connected'
        
        self.answer_label['text'] = 'Connecting...'
        threadObj = threading.Thread(target=connect_to_smtp)
        threadObj.start()

        
    def disconnect(self):
        """Disconnects from smtp server"""
        if self.smtpObj == 0:
            self.answer_label['text'] = 'Already disconnected'
        else:
            self.smtpObj.quit()
            self.answer_label['text'] = 'Disconnected'

   
    def send_mail(self):
        """Send the email"""
        
        def send_via_smtp():
            # Create an email message
            msg            = EmailMessage()
            msg['From']    = self.from_entry.get()
            msg['To']      = self.to_entry.get()
            msg['Subject'] = self.subject_entry.get()
            msg.set_content(self.msg_entry.get())
            
            # Send the message
            self.smtpObj.send_message(msg)
            self.answer_label['text'] = 'Sent'
            
        self.answer_label['text'] = 'Sending...'
        threadObj = threading.Thread(target=send_via_smtp)
        threadObj.start()
 
    def init_gui(self):
        """Builds GUI."""
        self.root.title('Vonk e-mail client')
        self.root.option_add('*tearOff', 'FALSE')
 
        self.grid(column=0, row=0, sticky='nsew')
        
        # Row 0
        ttk.Label(self, text='Compose an e-mail').grid(column=0, row=0, columnspan=4)
 
        # Row 1
        ttk.Separator(self, orient='horizontal').grid(column=0, row=1, columnspan=4, sticky='ew')
        
        # Row 2
        ttk.Label(self, text='From:').grid(column=0, row=2, sticky='w')
        self.from_entry = ttk.Entry(self, width=20)
        self.from_entry.grid(column=1, row=2)
        self.from_entry.insert(0, 'user@hotmail.com')
        
        # Row 3
        ttk.Label(self, text='Password:').grid(column=0, row=3, sticky='w')
        self.pswd_entry = ttk.Entry(self, width=20, show="*")
        self.pswd_entry.grid(column=1, row=3)
        
        # Row 4
        ttk.Label(self, text='SMTP server:').grid(column=0, row=4, sticky='w')
        self.smtp_entry = ttk.Entry(self, width=20)
        self.smtp_entry.grid(column=1, row=4)
        self.smtp_entry.insert(0, 'smtp-mail.outlook.com')
        
        # Row 5
        self.connect_button = ttk.Button(self, text='Connect', command=self.connect)
        self.connect_button.grid(column=0, row=5, columnspan=1)
        
        self.disconnect_button = ttk.Button(self, text='Disconnect', command=self.disconnect)
        self.disconnect_button.grid(column=1, row=5, columnspan=1)
        
        # Row 6
        ttk.Separator(self, orient='horizontal').grid(column=0, row=6, columnspan=4, sticky='ew')
        
        # Row 7
        ttk.Label(self, text='To:').grid(column=0, row=7, sticky='w')
        self.to_entry = ttk.Entry(self, width=20)
        self.to_entry.grid(column=1, row=7)
        self.to_entry.insert(0, 'user@hotmail.com')
        
        # Row 8
        ttk.Label(self, text='Subject:').grid(column=0, row=8, sticky='w')
        self.subject_entry = ttk.Entry(self, width=20)
        self.subject_entry.grid(column=1, row=8)
        self.subject_entry.insert(0, 'Test')
        
        # Row 9
        ttk.Label(self, text='Message:').grid(column=0, row=9, sticky='w')
        self.msg_entry = ttk.Entry(self, width=30)
        self.msg_entry.grid(column=1, row=9)
        self.msg_entry.insert(0, 'Does it work?')
        
        # Row 10
        self.calc_button = ttk.Button(self, text='Send e-mail', command=self.send_mail)
        self.calc_button.grid(column=0, row=10, columnspan=4)
        
        # New frame -  Status
        self.answer_frame = ttk.LabelFrame(self, text='Status',height=100)
        self.answer_frame.grid(column=0, row=11, columnspan=4, sticky='nesw')
 
        self.answer_label = ttk.Label(self.answer_frame, text='')
        self.answer_label.grid(column=0, row=0)

        for child in self.winfo_children():
            child.grid_configure(padx=5, pady=5)
 
if __name__ == '__main__':
    root = tkinter.Tk()
    EmailSender(root)
    root.mainloop()