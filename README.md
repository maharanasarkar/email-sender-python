# Email Sender Python

Welcome to the *Email Sender Python* repository! This versatile Python program is designed to simplify the process of sending emails to a targeted audience stored in a CSV file. Whether you're managing a mailing list, sending updates, or reaching out to clients, this tool provides a streamlined solution for automating the email dispatching process.

## Features

- **CSV Integration:** Easily import and manage your targeted audience using a CSV file. This allows you to maintain a structured and organized list of recipients.

- **Automation:** Save time and effort by automating the email dispatching process. This tool is designed to handle repetitive tasks, making it ideal for managing large mailing lists.

- **Versatility:** Whether you're sending newsletters, updates, or personalized messages, this Python program adapts to various use cases, providing a versatile solution for different email communication needs.

## How to Use

1. **Create a Virtual Environment:**
    ```bash
    # Create a new virtual environment
    python -m venv myenv
    
    # Activate the virtual environment
    source myenv/bin/activate      # For Linux/Mac
    .\myenv\Scripts\activate       # For Windows
    ```

2. **Install Dependencies:**
    ```bash
    # Install required dependencies
    pip install -r requirements.txt
    ```
3. **Setup ```.env``` file:**
    ```
    HOST=
    PORT=
    PASSWORD=
    EMAIL=
    ```

4. **Prepare Your CSV File:** Organize your recipient data in a CSV file, including necessary information such as email addresses, names, and any other relevant details.

5. **Configure Email Settings:** Customize the email settings within the program to align with your SMTP server details, authentication credentials, and other parameters.

6. **Run the Program:** Execute the Python script to initiate the email sending process. The program will read the CSV file, draft personalized emails, and dispatch them to the specified recipients.

7. **Monitor Progress:** Keep track of the email sending progress, receive status updates, and handle any potential issues that may arise during the process.


## Contribution Guidelines

We welcome contributions from the community to enhance the functionality and usability of this tool. If you have ideas for new features, improvements, or bug fixes, feel free to submit a pull request.

```Happy emailing! 📧✨```