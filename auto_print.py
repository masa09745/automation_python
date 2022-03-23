import win32api
import sys
import os
import time
import glob
from natsort import natsorted

import tkinter as tk
import tkinter.filedialog

def auto_print(path):
  win32api.shellExcute(0, "print", path, None, ".", 0)
  print ("Printed:" + path)
  

def filepath():
  dir = os.getcwd()
  root = tk.Tk()
  root.withdraw()

  target_dir = tkinter.filedialog.askdirectory(initialdir = dir)
  files = glob.glob(os.path.join(target_dir, '*.docx'))
  sort_list = natsorted(files)
  for filepath in sort_list:
    auto_print(filepath)
    time.sleep(3)
    

filepath ()
auto_print()

