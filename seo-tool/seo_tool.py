from tkinter import *
from tkinter import ttk
from tkinter import filedialog, messagebox
from threading import Thread
from queue import Queue
import requests
from scrapy import Selector
from multiprocessing import cpu_count
import time
import csv


class Root(Tk):
    def __init__(self):
        super(Root, self).__init__()
        self.title("Seo tool check")
        self.minsize(640, 400)
        self.wm_iconbitmap('icon.ico')

        self.labelFrame = ttk.LabelFrame(self, text="Open File")
        self.labelFrame.grid(column=0, row=1, padx=20, pady=20)

        self.ui()
        self.tree()

    def ui(self):
        self.button = ttk.Button(self.labelFrame, text="Browse A File", command=self.fileDialog)
        self.button.grid(column=1, row=1)
        self.count_threads = StringVar()
        self.delay = StringVar()
        self.custom_selector = StringVar()
        threads_label = Label(text='Count threads')
        threads_label.grid(column=1, row=1, sticky="n")
        count_threads = Entry(textvariable=self.count_threads)
        count_threads.grid(column=1, row=1)
        count_threads.insert(0, cpu_count())

        delay_label = Label(text='Delay (sec)')
        delay_label.grid(column=2, row=1, sticky="n")
        delay = Entry(textvariable=self.delay)
        delay.grid(column=2, row=1)
        delay.insert(0, 0.05)

        custom_selector_label = Label(text='Custom selector')
        custom_selector_label.grid(column=3, row=1, sticky="n")
        custom_selector_entry = Entry(textvariable=self.custom_selector)
        custom_selector_entry.grid(column=3, row=1)

        self.run_button = ttk.Button(self, text="Run", command=self.run)
        self.run_button.grid(row=6, column=0)

        self.delete_button = ttk.Button(self, text="Delete item", command=self.delete_data)
        self.delete_button.grid(row=6, column=1)

        self.run_button = ttk.Button(self, text="To csv", command=self.results_to_csv)
        self.run_button.grid(row=6, column=2)

    def tree(self):
        # Set the treeview
        self.colums = ('Link', 'Status', 'Title', 'H1', 'Description', 'Keywords', 'Sks', 'Custom selector')
        self.tree = ttk.Treeview(self, columns=self.colums)

        # Set the heading (Attribute Names)
        self.tree.heading('#0', text='Link')
        self.tree.heading('#1', text='Status')
        self.tree.heading('#2', text='Title')
        self.tree.heading('#3', text='H1')
        self.tree.heading('#4', text='Description')
        self.tree.heading('#5', text='Keywords')
        self.tree.heading('#6', text='Sks')
        self.tree.heading('#7', text='Custom selector')

        self.tree.column('#0', stretch=NO)
        self.tree.column('#1', stretch=NO, width=100)
        self.tree.column('#2', stretch=NO)
        self.tree.column('#3', stretch=NO)
        self.tree.column('#4', stretch=NO)
        self.tree.column('#5', stretch=NO, width=100)
        self.tree.column('#7', stretch=NO)

        self.tree.grid(row=5, columnspan=7, sticky='nsew')
        self.treeview = self.tree

        self.id = 0
        self.iid = 0

    def fileDialog(self):
        filename = filedialog.askopenfilename(initialdir="/", title="Select A File", filetypes=[("Csv files", ".csv")])
        self.label = ttk.Label(self.labelFrame, text="")
        self.label.grid(column=1, row=2)
        self.label.configure(text=filename)

        if filename:
            self.treeview.delete(*self.treeview.get_children())
            file = open(filename)
            reader = csv.reader(file)

            for item in reader:
                self.treeview.insert('', 'end', iid=self.iid, text=item[0])
                self.iid = self.iid + 1
                self.id = self.id + 1

    def delete_data(self):
        row_id = int(self.tree.focus())
        self.treeview.delete(row_id)

    def results_to_csv(self):
        if len(self.tree.get_children()) == 0:
            messagebox.showerror('Error', 'Run parsing before download file')
            return

        filename = filedialog.asksaveasfilename(initialdir="/", title="Save as...", filetypes=[("Csv files", ".csv")])

        if not filename.endswith('.csv'):
            filename = '{}.csv'.format(filename)

        with open(filename, 'w') as file:
            writer = csv.DictWriter(file, fieldnames=list(self.colums))
            writer.writeheader()
            for children in self.tree.get_children():
                data = self.tree.item(children)

                row = {
                    'Link': data['text'],
                    'Status': data['values'][0] if len(data['values']) >= 1 else '',
                    'Title': data['values'][1] if len(data['values']) >= 2 else '',
                    'H1': data['values'][2] if len(data['values']) >= 3 else '',
                    'Description': data['values'][3] if len(data['values']) >= 4 else '',
                    'Keywords': data['values'][4] if len(data['values']) >= 5 else '',
                    'Sks': data['values'][5] if len(data['values']) >= 6 else '',
                    'Custom selector': data['values'][6] if len(data['values']) >= 7 else '',
                }
                writer.writerow(row)

    def progressbar(self, length):
        self.progress_bar = ttk.Progressbar(self, orient=HORIZONTAL, length=length, mode='determinate')
        self.progress_bar.grid(row=7, column=0)
        self.progress_bar.start()
        self.progress_bar.update()

    def run(self):
        self.queue = Queue()
        for children in self.tree.get_children():
            self.queue.put(children)

        threads = []
        print('Debug run parsing with {} count threads'.format(self.count_threads.get()))
        print('Debug run parsing with {} delay seconds'.format(self.delay.get()))
        for i in range(int(self.count_threads.get())):
            t = Thread(target=self.run_processing)
            t.start()
            threads.append(t)

    def run_processing(self):
        while True:
            time.sleep(float(self.delay.get()))
            if self.queue.empty():
                return True

            task = self.queue.get()
            self.process_task(task)

    def process_task(self, task_id):
        task = self.treeview.item(task_id)
        url = task['text']
        response = requests.get(url)
        html = response.content
        selector = Selector(text=html)
        title = selector.xpath('//title/text()').extract_first()
        h1 = selector.xpath('//h1/text()').extract_first()
        description = selector.xpath("//meta[@name='description']/@content").extract_first()
        keywords = selector.xpath("//meta[@name='keywords']/@content").extract_first()
        sks = ','.join(selector.css('span[data-sks-key]::text').extract())
        self.treeview.item(task_id, values=(response.status_code, title, h1, description, keywords, sks))


root = Root()
root.mainloop()
