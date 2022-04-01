import os
import sys
import time
import shutil
import pandas as pd
from pathlib import Path
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

path = input("""
Please enter the directory you want to monitor, if not add the directory it will take the directory where the script is running
""")

pathGeneral = os.getcwd()
if len(path) == 0:
    path = pathGeneral + "/Files"
    if not os.path.isdir(path):
        os.mkdir(path)

    print("The default directory is used")

if not os.path.isdir(path):
    ValueError("Directory not found")
    sys.exit(0)

process = pathGeneral + "/Processed" 
not_process = pathGeneral + "/Not applicable"
if not os.path.isdir(process):
    os.mkdir(process)

if not os.path.isdir(not_process):
    os.mkdir(not_process)

class Watcher:
    directory_to_watch = path
    observer = Observer()
    
    @classmethod
    def run(cls):
        event_handler = Handler()
        cls.observer.schedule(event_handler, cls.directory_to_watch, recursive=True)
        cls.observer.start()
        try:
            while True:
                time.sleep(5)
        except:
            cls.observer.stop()
            print("Error")

        cls.observer.join()

class Handler(FileSystemEventHandler):

    @staticmethod
    def on_any_event(event):
        if event.is_directory:
            return None

        if event.event_type == 'created' and Path(event.src_path).suffix in [".xlsx", ".xls"]:
            print(event.src_path)
            file_detect = pd.ExcelFile(event.src_path)
            master = pd.ExcelFile("MasterBook.xlsx")
            with pd.ExcelWriter("MasterBook.xlsx", engine='xlsxwriter') as writer:
                for sheet_name in master.sheet_names:
                    sheet = master.parse(sheet_name)
                    sheet.to_excel(writer, sheet_name=sheet_name)
                
                for sheet_name in file_detect.sheet_names:
                    sheet = file_detect.parse(sheet_name)
                    sheet.to_excel(writer, sheet_name=sheet_name)

            shutil.copyfile(event.src_path, process + f"/{sheet_name}_Copy{Path(event.src_path).suffix}")
            #os.remove(event.src_path)
            print(f"File {event.src_path} moved to {process}")

        if event.event_type == 'created' and Path(event.src_path).suffix not in [".xlsx", ".xls"]:
            filename = event.src_path.split("/")[-1]
            shutil.copyfile(event.src_path, not_process + f"/{filename}")
            #os.remove(event.src_path)
            print(f"File {event.src_path} moved to {process}")

if __name__ == "__main__":
    Watcher.run()
