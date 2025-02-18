import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import threading, queue

from id_generator import *


class GUI():
    def __init__(self, root):
        
        root.title("Vedion ID-generator ver.1.0.0225")
        root.geometry("500x260")
        root.resizable(False, False)
        
        root.grid_columnconfigure(0, weight=0)
        root.grid_columnconfigure(1, weight=1)
        
        self.file_paths = [None] * 3
        
        self.config_filename = "config.cfg"
        
        self.hdd_switch = tk.BooleanVar()
        
        # Open file buttons
        self.file1_button = ttk.Button(root, text="Odśwież", command=self.file1_button_callback)
        self.file2_button = ttk.Button(root, text="...", command=lambda: self.open_file(1))
        self.file3_button = ttk.Button(root, text="...", command=self.open_directory)
        self.hdd_switch_checkbox = ttk.Checkbutton(root, text="ID z dyskami", variable=self.hdd_switch, command=self.hdd_switch_callback)
        
        # Entries with file names
        self.entry1 = ttk.Entry(root, state="disabled")
        self.entry2 = ttk.Entry(root, state="disabled")
        self.entry3 = ttk.Entry(root, state="disabled")
        
        self.label1 = ttk.Label(root, text="Link do pliku M2 M47:")
        self.label2 = ttk.Label(root, text="Plik z raportem:")
        self.label3 = ttk.Label(root, text="Folder wyjściowy:")
        
        # Button that start thread that searches for ID's and progressbar showing progress of the process
        self.search_button = ttk.Button(root, text="Wyszukaj", command=self.search_button_callback)
        self.progressbar = ttk.Progressbar(root, orient="horizontal", length=120, mode="indeterminate")
        
        self.entry1.grid(row=0, column=1, padx=(10,2), pady=(20, 1), sticky="we")
        self.entry2.grid(row=1, column=1, padx=(10,2), pady=1, sticky="we")
        self.entry3.grid(row=2, column=1, padx=(10,2), pady=1, sticky="we")
        
        self.label1.grid(row=0, column=0, padx=(10,2), pady=(20, 1), sticky="we")
        self.label2.grid(row=1, column=0, padx=(10,2), pady=1, sticky="we")
        self.label3.grid(row=2, column=0, padx=(10,2), pady=1, sticky="we")
        
        self.hdd_switch_checkbox.grid(row=3, column=1, padx=(0,40), pady=10)
        
        # Placeholder texts
        self.fill_entry(self.entry1, "")
        self.entry1.configure(state='normal')
        self.fill_entry(self.entry2, "")
        self.fill_entry(self.entry3, "")
        
        self.read_config()
        
        self.file1_button.grid(row=0, column=2, padx=(0,10), pady=(20, 1), sticky="w")
        self.file2_button.grid(row=1, column=2, padx=(0,10), pady=1, sticky="w")
        self.file3_button.grid(row=2, column=2, padx=(0,10), pady=1, sticky="w")
        
        self.search_button.grid(row=4, column=1, padx=(0,40), pady=(0, 30))
        self.progressbar.grid(row=5, column=1, padx=(0,40), pady=(10, 2))
        
        self.label_status = ttk.Label(root, text="")
        self.label_status.grid(row=6, column=1, padx=(0,40))
        
        # Init IDGenerator
        self.generator = IDGenerator()


    def file1_button_callback(self):
        self.generator.toggle_fID_changed()

    # Reads entry contents from a config
    def read_config(self):
        try:
            with open(self.config_filename, "r") as file:
                lines = file.readlines()
                if lines:
                    try:
                        self.fill_entry(self.entry1, lines[0].rstrip("\n")) # remove "\n" from a end of a string
                        self.entry1.configure(state='normal')
                        self.fill_entry(self.entry3, lines[1])
                        self.file_paths[2] = lines[1]
                    except:
                        pass
        except:
            with open(self.config_filename, "a+") as file:  # Create file if it does not exist
                file.read()
              
    # Writes entry contents to a config
    def write_config(self):
        with open(self.config_filename, "w") as file:
            entry1_text = self.entry1.get()
            entry3_text = self.entry3.get()
            lines = [entry1_text, '\n', entry3_text]
            file.writelines(lines)

            
    # Reads file path from input file
    def open_file(self, button_id):
        file_path = filedialog.askopenfilename(
            title="Wybierz plik",
            filetypes=(("MS Excel files", "*.xls *xlsx"), ("All files", "*.*"))
        )
        if file_path:
            self.file_paths[button_id] = file_path
            if button_id == 1:
                self.fill_entry2(file_path)
        else:
            # messagebox.showinfo("Nie wybrano pliku", "Nie wybrano pliku.")
            pass
    
    def open_directory(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.fill_entry3(folder_path)
            self.file_paths[2] = folder_path
            
    def search_button_callback(self):
        # Get link to google spreadsheet from enrtry1
        spreasheet_link = self.entry1.get()
        # Open and read files then this function creates thread that search for ID's
        self.read_files_thread(self.generator, spreasheet_link, self.file_paths[1], self.file_paths[2])
        
    
    def hdd_switch_callback(self):
        print(self.hdd_switch.get())
        self.generator.toggle_hdd_switch(self.hdd_switch.get())
        
    # Makes a thread wchich runs functions from IDGenerator that search for ID's
    # It is non-blocking, queue with returned value is checked every 50ms
    def search_thread(self, generator):
        generator_queue = queue.Queue() # return value queue
        thread = threading.Thread(target=generator.process_files, args=(generator_queue,))
        
        # Starts the thread and progressbar
        thread.start()
        self.label_status.configure(text="Wyszukiwanie ID")
        self.progressbar.start(10)
        
        # Checks queue for returned vales every 50ms
        root.after(50, self.check_result_search, generator_queue)
        
    
    def read_files_thread(self, generator, spreasheet_link, file_path1, file_path2):
        generator_queue = queue.Queue() # return value queue
        thread = threading.Thread(target=generator.read_files, args=(generator_queue, spreasheet_link, file_path1, file_path2,))
        thread.daemon = False
        
        # Starts the thread and progressbar
        thread.start()
        self.label_status.configure(text="Pobieranie ID")
        self.progressbar.start(10)
        
        # Checks queue for returned vales every 50ms
        root.after(50, self.check_result_read, generator_queue)
    

    # Function to check and retrieve the result from the queue
    def check_result_search(self, result_queue):
        try:
            err = result_queue.get_nowait()  # Try to get the result and update 
            # If returned value is int then function finished properly and found ID's
            # Else there were error
            if isinstance(err, int):    
                self.progressbar.stop()
                self.label_status.configure(text="")
                msg = "Znaleziono " + str(err) + " ID"
                messagebox.showinfo("Znaleziono", msg)  # Show how many ID's were found
                return err
            else:   
                self.progressbar.stop()
                self.label_status.configure(text="")
                messagebox.showinfo("Error!", err)  #Show errors
                return err
        except queue.Empty:
            root.after(50, self.check_result_search, result_queue)  # Check again later if no result yet
            
    
    # Function to check and retrieve the result from the queue
    def check_result_read(self, result_queue):
        try:
            err = result_queue.get_nowait()  # Try to get the result and update 
            # If returned value is int then function finished properly and found ID's
            # Else there were error
            if err is None:    
                self.progressbar.stop()
                self.label_status.configure(text="")
                # Finds ID's
                self.search_thread(self.generator)
                return err
            else:   
                self.progressbar.stop()
                self.label_status.configure(text="")
                messagebox.showinfo("Error!", err)  #Show errors
                return err
        except queue.Empty:
            root.after(50, self.check_result_read, result_queue)  # Check again later if no result yet


    # Fill given entry end entry handle
    def fill_entry(self, entry, text):
        entry.configure(state='normal')
        entry.delete(0, tk.END)
        entry.insert(0, text)
        entry.configure(state='disabled')
    
    
    # Fill entry callbacks        
    def fill_entry1(self, text):
        self.fill_entry(self.entry1, text)
    
    
    def fill_entry2(self, text):
        self.fill_entry(self.entry2, text)
        
        
    def fill_entry3(self, text):
        self.fill_entry(self.entry3, text)
        
        
    def on_close(self):
        self.write_config()
        root.destroy()


root = tk.Tk()
gui = GUI(root)
root.protocol("WM_DELETE_WINDOW", gui.on_close)
root.mainloop()