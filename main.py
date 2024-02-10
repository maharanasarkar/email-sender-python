import os
import csv
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from dotenv import load_dotenv
from tkinter import Tk, Label, Entry, Button, filedialog, Frame, Text, Scrollbar, ttk

load_dotenv()

class EmailSenderApp:
    def __init__(self, master):
        self.master = master
        master.title("Email Sender App")

        # Styling
        self.master.configure(bg="#f0f0f0")
        self.master.geometry("800x500")

        # Header Frame
        header_frame = Frame(master, bg="#f0f0f0")
        header_frame.grid(row=0, column=0, columnspan=3, sticky="nsew")

        # Content Frame
        content_frame = Frame(master, bg="#f0f0f0")
        content_frame.grid(row=1, column=0, columnspan=3, sticky="nsew")

        self.label_csv = Label(header_frame, text="CSV or Excel File:", bg="#f0f0f0")
        self.label_csv.grid(row=0, column=0, padx=10, pady=10, sticky="w")

        self.entry_csv = Entry(header_frame, width=40)
        self.entry_csv.grid(row=0, column=1, padx=10, pady=10)

        self.button_browse = Button(header_frame, text="Browse", command=self.browse_file, bg="#007BFF", fg="white")
        self.button_browse.grid(row=0, column=2, padx=10, pady=10, sticky="w")

        self.label_subject = Label(content_frame, text="Subject:", bg="#f0f0f0")
        self.label_subject.grid(row=0, column=0, padx=10, pady=5, sticky="w")

        self.entry_subject = Entry(content_frame, width=40)
        self.entry_subject.grid(row=0, column=1, padx=10, pady=5)

        self.label_message = Label(content_frame, text="Message Body:", bg="#f0f0f0")
        self.label_message.grid(row=1, column=0, padx=10, pady=5, sticky="w")

        self.entry_message = Text(content_frame, width=40, height=10)
        self.entry_message.grid(row=1, column=1, padx=10, pady=5)

        self.button_send = Button(content_frame, text="Send Emails", command=self.send_emails, bg="#28a745", fg="white")
        self.button_send.grid(row=1, column=3, pady=(20, 0))

        # Manual Emails Frame
        manual_frame = Frame(master, bg="#f0f0f0")
        manual_frame.grid(row=2, column=0, columnspan=3, sticky="nsew")

        self.label_manual_emails = Label(manual_frame, text="Manual Emails (comma-separated):", bg="#f0f0f0")
        self.label_manual_emails.grid(row=0, column=0, padx=10, pady=5, sticky="w")

        self.entry_manual_emails = Entry(manual_frame, width=40)
        self.entry_manual_emails.grid(row=0, column=1, padx=10, pady=5)

        self.button_add_manual = Button(manual_frame, text="Add Manual Emails", command=self.add_manual_emails, bg="#007BFF", fg="white")
        self.button_add_manual.grid(row=0, column=2, padx=10, pady=5, sticky="w")

        # Emails Display Frame
        display_frame = Frame(master, bg="#f0f0f0")
        display_frame.grid(row=3, column=0, columnspan=3, sticky="nsew")

        self.emails_text = Text(display_frame, wrap="word", width=60, height=10)
        self.emails_text.grid(row=0, column=0, padx=10, pady=10)

        self.scrollbar = Scrollbar(display_frame, command=self.emails_text.yview)
        self.scrollbar.grid(row=0, column=1, sticky="nsew")
        self.emails_text.config(yscrollcommand=self.scrollbar.set)

        # Configure column and row weights for resizing
        for i in range(3):
            master.columnconfigure(i, weight=1)
        for i in range(4):
            master.rowconfigure(i, weight=1)

    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx")])
        self.entry_csv.delete(0, 'end')
        self.entry_csv.insert(0, file_path)
        self.display_emails()

    def send_emails(self):
        email = os.getenv("EMAIL")
        password = os.getenv("PASSWORD")

        csv_file = self.entry_csv.get()
        subject = self.entry_subject.get()
        message_body = self.entry_message.get()

        try:
            connection = smtplib.SMTP(host=os.getenv("HOST"), port=os.getenv("PORT"))
            connection.starttls()
            connection.login(user=email, password=password)

            msg = MIMEMultipart()
            msg['Subject'] = subject
            msg.attach(MIMEText(message_body, 'html'))

            if csv_file.endswith('.csv'):
                with open(csv_file, 'r') as file:
                    reader = csv.reader(file)
                    next(reader)  # Skip the header row if it exists
                    for row in reader:
                        target_email = row[0]
                        connection.sendmail(from_addr=email, to_addrs=target_email, msg=msg.as_string())

            elif csv_file.endswith('.xlsx'):
                df = pd.read_excel(csv_file)
                for index, row in df.iterrows():
                    target_email = row['EmailColumn']  # Replace 'EmailColumn' with the actual column name containing emails
                    connection.sendmail(from_addr=email, to_addrs=target_email, msg=msg.as_string())

            else:
                raise ValueError("Unsupported file format. Please select a CSV or Excel file.")

        except Exception as e:
            print(f"An error occurred: {e}")

        finally:
            if connection:
                connection.quit()
        self.display_emails()


    def display_emails(self):
        # Display extracted emails in the Text widget
        csv_file = self.entry_csv.get()
        if csv_file.endswith('.csv'):
            with open(csv_file, 'r') as file:
                reader = csv.reader(file)
                next(reader)  # Skip the header row if it exists
                extracted_emails = [row[0] for row in reader]

        elif csv_file.endswith('.xlsx'):
            df = pd.read_excel(csv_file)
            extracted_emails = df['EmailColumn'].tolist()  # Replace 'EmailColumn' with the actual column name containing emails

        else:
            extracted_emails = []

        # Clear previous content
        self.emails_text.delete('1.0', 'end')

        # Display emails
        for email in extracted_emails:
            self.emails_text.insert('end', email + '\n')

    def add_manual_emails(self):
        # Add manually entered emails to the existing ones
        manual_emails = self.entry_manual_emails.get()
        if manual_emails:
            manual_email_list = [email.strip() for email in manual_emails.split(',')]
            for email in manual_email_list:
                self.emails_text.insert('end', email + '\n')




if __name__ == "__main__":
    root = Tk()
    app = EmailSenderApp(root)
    root.mainloop()
