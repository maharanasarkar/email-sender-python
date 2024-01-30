import os
import csv
import pandas as pd
import smtplib
from dotenv import load_dotenv
from tkinter import Tk, Label, Entry, Button, filedialog

load_dotenv()

class EmailSenderApp:
    def __init__(self, master):
        self.master = master
        master.title("Email Sender App")

        self.label_csv = Label(master, text="CSV or Excel File:")
        self.label_csv.grid(row=0, column=0, padx=10, pady=10)

        self.entry_csv = Entry(master)
        self.entry_csv.grid(row=0, column=1, padx=10, pady=10)

        self.button_browse = Button(master, text="Browse", command=self.browse_file)
        self.button_browse.grid(row=0, column=2, padx=10, pady=10)

        self.label_subject = Label(master, text="Subject:")
        self.label_subject.grid(row=1, column=0, padx=10, pady=10)

        self.entry_subject = Entry(master)
        self.entry_subject.grid(row=1, column=1, padx=10, pady=10)

        self.label_message = Label(master, text="Message Body:")
        self.label_message.grid(row=2, column=0, padx=10, pady=10)

        self.entry_message = Entry(master)
        self.entry_message.grid(row=2, column=1, padx=10, pady=10)

        self.button_send = Button(master, text="Send Emails", command=self.send_emails)
        self.button_send.grid(row=3, column=1, pady=20)

    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx")])
        self.entry_csv.delete(0, 'end')
        self.entry_csv.insert(0, file_path)

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

            if csv_file.endswith('.csv'):
                with open(csv_file, 'r') as file:
                    reader = csv.reader(file)
                    next(reader)  # Skip the header row if it exists
                    for row in reader:
                        target_email = row[0]
                        connection.sendmail(from_addr=email, to_addrs=target_email, msg=f"Subject: {subject}\n\n{message_body}")

            elif csv_file.endswith('.xlsx'):
                df = pd.read_excel(csv_file)
                for index, row in df.iterrows():
                    target_email = row['EmailColumn']  # Replace 'EmailColumn' with the actual column name containing emails
                    connection.sendmail(from_addr=email, to_addrs=target_email, msg=f"Subject: {subject}\n\n{message_body}")

            else:
                raise ValueError("Unsupported file format. Please select a CSV or Excel file.")

        except Exception as e:
            print(f"An error occurred: {e}")

        finally:
            if connection:
                connection.quit()


if __name__ == "__main__":
    root = Tk()
    app = EmailSenderApp(root)
    root.mainloop()
