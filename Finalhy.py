#There we import the modules and libraries  that are required
from tkinter import Tk, Label, Entry, Button, StringVar, Text, Scrollbar, messagebox, filedialog
from tkinter import ttk  # Import ttk module
from selectorlib import Extractor
import requests
import json
import webbrowser
import csv
import pandas as pd
#This is Title name of program
class ProductScraperGUI:
    def __init__(self, master):
        self.master = master
        master.title(" Product Search")
        # Create a style object
        self.style = ttk.Style()
        # Configure the style with some basic settings
        self.style.configure("TLabel", font=("Arial", 12),width=16)
        self.style.configure("TEntry", font=("Arial", 12),width=16)
        self.style.configure("TButton", font=("Arial", 12),width=16)
# making the labels
        self.label = ttk.Label(master, text="Enter search text:", style="TLabel")  # Apply the style
        self.label.pack()

        self.entry_var = StringVar()
        self.entry = ttk.Entry(master, textvariable=self.entry_var, width=50, style="TEntry")  # Apply the style
        self.entry.pack()
#making buttons
        self.scrape_button = ttk.Button(master, text="Scrape", command=self.scrape_and_display, style="TButton")  # Apply the style
        self.scrape_button.pack(pady=10)
#making  output text
        self.output_text = Text(master, wrap="word", height=20, width=80)
        self.output_text.pack(pady=10)
#making scroll bar
        self.scrollbar = Scrollbar(master, command=self.output_text.yview)
        self.scrollbar.pack(side="right", fill="y")
        self.output_text.config(yscrollcommand=self.scrollbar.set)
#making CVS button
        self.save_csv_button = ttk.Button(master, text="Save as CSV", command=self.save_as_csv, style="TButton")  # Apply the style
        self.save_csv_button.pack(pady=5)


#making the directory button
        self.choose_directory_button = ttk.Button(master, text="Choose Directory", command=self.choose_directory, style="TButton")  # Apply the style
        self.choose_directory_button.pack(pady=5)

        self.selected_directory = ""
#in this we fix the url of amazon and take input from user
    def scrape_and_display(self):
        search_text = self.entry_var.get()
        url = "https://www.amazon.com/s?k=" + search_text

        data = self.scrape(url)
# searching the data from the website like product title,price,url,description
        if data:
            self.output_text.delete(1.0, "end")
            for product in data['products']:
                self.output_text.insert("end", f"Title: {product.get('title', '')}\n")
                self.output_text.insert("end", f"Price: {product.get('price', '')}\n")
                self.output_text.insert("end", f"Description: {product.get('description', '')}\n")
                self.output_text.insert("end", f"Search URL: {url}\n")
                self.output_text.insert("end", "\n")
#This function is for Browser Support
    def scrape(self, url):
        headers = {
            'dnt': '1',
            'upgrade-insecure-requests': '1',
            'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_4) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.61 Safari/537.36',
            'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
            'sec-fetch-site': 'same-origin',
            'sec-fetch-mode': 'navigate',
            'sec-fetch-user': '?1',
            'sec-fetch-dest': 'document',
            'referer': 'https://www.amazon.com/',
            'accept-language': 'en-GB,en-US;q=0.9,en;q=0.8',
        }
# This is request to browser
        r = requests.get(url, headers=headers)
# if it not search or come error then it show that
        if r.status_code > 500:
            if "To discuss automated access to Amazon data please contact" in r.text:
                return None
            else:
                return None
#In this we extract the yaml file because it contain the style of searching
        e = Extractor.from_yaml_file('search_results.yml')
        return e.extract(r.text)
#This function is for Directory
    def choose_directory(self):
        self.selected_directory = filedialog.askdirectory()
        messagebox.showinfo("Directory Selected", f"Selected directory: {self.selected_directory}")
#this function  is for save file in Excel and extensions Cvs
    def save_as_csv(self):
        try:
            if not self.selected_directory:
                self.choose_directory()
            if self.selected_directory:
                data = self.get_scraped_data()
                if data:
                    df = pd.DataFrame(data['products'])
                    file_path = f"{self.selected_directory}/scraped_data.csv"
                    df.to_csv(file_path, index=False)
                    messagebox.showinfo("Save Success", f"Data saved to {file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Error saving data to CSV: {str(e)}")
#this function  is for save file in Excel and extensions xls
    def save_as_xls(self):
        try:
            if not self.selected_directory:
                self.choose_directory()
            if self.selected_directory:
                data = self.get_scraped_data()
                if data:
                    df = pd.DataFrame(data['products'])
                    file_path = f"{self.selected_directory}/scraped_data.xls"
                    df.to_excel(file_path, index=False)
                    messagebox.showinfo("Save Success", f"Data saved to {file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Error saving data to XLS: {str(e)}")
# this function is for taking input and add it with url
    def get_scraped_data(self):
        search_text = self.entry_var.get()
        url = "https://www.amazon.com/s?k=" + search_text
        return self.scrape(url)

# This is to Run GUI
if __name__ == "__main__":
    root = Tk()
    app = ProductScraperGUI(root)
    root.mainloop()
