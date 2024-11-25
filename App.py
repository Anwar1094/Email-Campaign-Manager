# Before trying to run this file kindly download some image files from: https://github.com/Anwar1094/Email-Campaign-Manager
# Import Necessary Libraries
from customtkinter import *
from tkinter import *
from tkinter import messagebox as msgbx
from PIL import Image
import time
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import threading

# Defining width and height of canvas
width = 200
height = 200

# Main App Class
class App(CTk):
    # Setting up server
    Server, port = 'smtp.gmail.com', 587
    # Creating Necesary Variables
    start_time = 0
    success_count = 0
    fail_count = 0
    err_msgs = []
    total_recipients = 1
    # Constructor
    def __init__(self):
        super().__init__()
        self.geometry(f'{width}x{height}+{self.winfo_screenwidth()//2+50}+{self.winfo_screenheight()//2-100}')
        self.resizable(0,0)
        self.title('Email Compaign Manager')
        self.wm_iconbitmap("email_services.ico")
        self.attributes('-top', True)
        self.overrideredirect(True)
        self.Logo = CTkLabel(self, text='', image=CTkImage(Image.open('email_services.jpg'), size=(200, 200)))
        self.Logo.place(x=0, y=0)

        # Variables # Login Credentail
        self.Sender_email_var = StringVar()
        self.password_var = StringVar()
        self.subject_var = StringVar()
        self.recepient_emails_var = StringVar()
        self.num_emails_var = StringVar(value='1')

    # Method to show buttons 
    def showOptions(self):
        CTkLabel(self, text='', bg_color='white', image=CTkImage(Image.open('EmailLogo.png'), size=(450, 150))).place(x=580, y=0)
        self.single_R_Btn = CTkButton(self, 40, 100, 80, text='', fg_color='white', bg_color='white', hover=False, image=CTkImage(Image.open('SingleRBtn.png'), size=(200, 100)), command=lambda: self.collectData('Single'))
        self.multi_R_Btn = CTkButton(self, 40, 100, 80, text='', fg_color='white', bg_color='white', hover=False, image=CTkImage(Image.open('MultiRBtn.png'), size=(200, 100)), command=lambda: self.collectData('Multi'))
        self.single_R_Btn.place(x=500, y=250)
        self.multi_R_Btn.place(x=780, y=250)

    # Method to go home window
    def Home(self):
        try:
            self.Frame.destroy()
            self.LogText.destroy()
        except:pass
        self.showOptions()

    # Method for creating gui of email sender based on button clicked
    def collectData(self, type):
        self.single_R_Btn.place_forget()
        self.multi_R_Btn.place_forget()
        self.Frame = CTkFrame(self, 1000, 400, bg_color='white', fg_color='#81d4fa', border_color='#394fb2', border_width=2)
        self.Frame.place(x=250, y=150)
        CTkLabel(self.Frame, text='To: ', text_color='black').place(x=20, y=20)
        CTkEntry(self.Frame, width=200, fg_color='white', text_color='black', placeholder_text_color='black', corner_radius=10, textvariable=self.recepient_emails_var, placeholder_text='Recepient Email').place(x=45, y=20)
        if type == 'Single':
            CTkLabel(self.Frame, text='No. of Mails: ', text_color='black').place(x=700, y=20)
            CTkEntry(self.Frame, width=200, fg_color='white', text_color='black', placeholder_text_color='black', corner_radius=10, textvariable=self.num_emails_var, placeholder_text='No. of mails', show='').place(x=780, y=20)
        CTkLabel(self.Frame, text='From: ', text_color='black').place(x=20, y=50)
        CTkEntry(self.Frame, width=200, fg_color='white', text_color='black', placeholder_text_color='black', corner_radius=10, textvariable=self.Sender_email_var, placeholder_text='Sender Email').place(x=65, y=50)
        CTkLabel(self.Frame, text='Password: ', text_color='black').place(x=700, y=50)
        CTkEntry(self.Frame, width=200, fg_color='white', text_color='black', placeholder_text_color='black', corner_radius=10, textvariable=self.password_var, placeholder_text='Password', show='*').place(x=780, y=50)
        CTkLabel(self.Frame, text='Subject: ', text_color='black').place(x=20, y=80)
        CTkEntry(self.Frame, width=200, fg_color='white', text_color='black', placeholder_text_color='black', corner_radius=10, textvariable=self.subject_var, placeholder_text='Subject').place(x=80, y=80)
        CTkLabel(self.Frame, text='Start Writing Your Mail...', text_color='black').place(x=20, y=110)
        self.msg = CTkTextbox(self.Frame, width=950, height=200, fg_color='white', text_color='black', corner_radius=10)
        self.msg.place(x=20, y=150)
        CTkButton(self.Frame, text='Send', fg_color='#1b61d1', width=50, corner_radius=50, text_color='white', command=lambda:threading.Thread(target=lambda: self.SendData(type)).start()).place(x=20, y=360)
        CTkButton(self.Frame, text='Home', fg_color='#1b61d1', width=50, corner_radius=50, text_color='white', command=lambda:self.Home()).place(x=900, y=360)

    # Creating Text box for output display
    def createTextBox(self):
        self.LogText = CTkTextbox(self, width=1000, height=200, corner_radius=10, bg_color='white', fg_color='black', text_color='white')
        self.LogText.place(x=250, y=560)
    
    # Method to create a template mail message
    def create_msg(self, sender, recipient, sub, msg):
        message = MIMEMultipart()
        message['From'] = sender
        message['To'] = recipient
        message['Subject'] = sub
        message.attach(MIMEText(msg, 'plain'))
        return message
    
    # Method to update text of Text Box
    def updateTextBox(self, text):
        self.LogText.insert('end', text)

    # Method to send data for sending emails
    def SendData(self, type):
        threading.Thread(target=lambda: self.createTextBox()).start()
        self.SenderEmail = self.Sender_email_var.get()
        self.Subject = self.subject_var.get()
        self.password = self.password_var.get()
        self.Msg = self.msg.get('0.0', 'end')
        self.singleSend = False
        if type == 'Single':
            self.RecepientEmails = [self.recepient_emails_var.get()]
            for i in range(int(self.num_emails_var.get())):
                self.SendMail(self.SenderEmail, self.RecepientEmails, self.Subject, self.Msg)
                self.singleSend = True
        elif type == 'Multi':
            self.RecepientEmails = self.recepient_emails_var.get().replace(' ', '').strip().split(',')
            self.total_recipients = len(self.RecepientEmails)
            self.SendMail(self.SenderEmail, self.RecepientEmails, self.Subject, self.Msg)
        self.report()

    # Method for sending emails
    def SendMail(self, sender, recepients, sub, msg):
        self.start_time = time.time()
        try:
            with smtplib.SMTP(self.Server, self.port) as server:
                server.starttls()
                server.login(sender, self.password)
                if self.singleSend != True:
                    threading.Thread(target=lambda: self.updateTextBox(f"Connected to {server}. Sending emails...\n")).start()
                for recipient in recepients:
                    try:
                        server.sendmail(from_addr=sender, to_addrs=recipient, msg=self.create_msg(sender, recipient, sub, msg).as_string())
                        self.success_count += 1
                        threading.Thread(target=lambda: self.updateTextBox(f'Email sent to: {recipient}\n')).start()
                    except Exception as e:
                        self.fail_count += 1
                        self.err_msgs.append(f"Failed to send email to {recipient}: {str(e)}\n")
                        threading.Thread(target=lambda:self.updateTextBox(f"Failed to send email to {recipient}: {e}\n")).start()
        except Exception as e:
            threading.Thread(target=lambda: self.updateTextBox(f'Error connecting to SMTP server: {e}\n')).start()

    # Method for creating summary report of mail sent
    def report(self):
        end_time = time.time()
        total_time = end_time - self.start_time
        threading.Thread(target=lambda: self.updateTextBox(f"\n--- Performance Report ---\n")).start()
        threading.Thread(target=lambda: self.updateTextBox(f"Total emails to send: {self.total_recipients}\n")).start()
        threading.Thread(target=lambda: self.updateTextBox(f"Emails sent successfully: {self.success_count}\n")).start()
        threading.Thread(target=lambda: self.updateTextBox(f"Emails failed: {self.fail_count}\n")).start()
        threading.Thread(target=lambda: self.updateTextBox(f"Time taken: {total_time:.2f} seconds\n")).start()
        if self.err_msgs:
            threading.Thread(target=lambda: self.updateTextBox("\n--- Error Messages ---\n")).start()
            for msg in self.err_msgs:
                threading.Thread(target=lambda: self.updateTextBox(f'{msg}')).start()

# Main class for startup welcome screen
class Main:
    def __init__(self, root):
        self.root = root
        root.Logo.place_forget()
        root.geometry(f'{width}x{height}+{root.winfo_screenwidth()//2-100}+{root.winfo_screenheight()//2-200}')
        root.overrideredirect(False)
        root.attributes('-top', False)
        root.config(background='white')
        root.state('zoomed')

# Calling Method and classes
if __name__ == '__main__':
    app = App()
    main = None
    app.after(1000, lambda: Main(app))
    app.after(1000, lambda: app.showOptions())
    app.mainloop()
