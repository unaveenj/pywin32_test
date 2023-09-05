import win32com.client
import tkinter as tk
from tkinter import ttk
import re


def wildcard_to_regex(pattern):
    """Convert a wildcard pattern to regex pattern."""
    return "^" + pattern.replace("*", ".*").replace("?", ".") + "$"


def count_outlook_emails_with_subject_regex(subject_pattern):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)

    # Broad filter to fetch all emails
    all_emails = inbox.Items

    # Use regex to filter these emails
    regex = re.compile(subject_pattern, re.IGNORECASE)
    matching_emails = [email for email in all_emails if regex.match(email.Subject)]

    return len(matching_emails)


def on_check_emails():
    user_input = pattern_entry.get()
    subject_to_search = wildcard_to_regex(user_input)
    count = count_outlook_emails_with_subject_regex(subject_to_search)
    result_label.config(text=f"Number of emails with the subject pattern: {count}")



root = tk.Tk()
root.title("Email Count Checker")


ttk.Label(root, text="Enter pattern (use * and ? as wildcards):").pack(pady=10, padx=10)
pattern_entry = ttk.Entry(root, width=40)
pattern_entry.pack(pady=10, padx=10)
pattern_entry.insert(0, "*Position Type: Faculty Internships*")  # Default pattern


check_button = ttk.Button(root, text="Check Emails", command=on_check_emails)
check_button.pack(pady=20)


result_label = ttk.Label(root, text="")
result_label.pack(pady=20)

root.mainloop()
