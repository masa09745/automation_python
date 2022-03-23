#import win32api
import sys
import os
import time
import glob
from natsort import natsorted

import tkinter as tk
import tkinter.filedialog

def filepath():
  dir = os.getcwd()
  root = tk.Tk()
  root.withdraw()

  target_dir = tkinter.filedialog.askdirectory(initialdir = dir)
  files = glob.glob(os.path.join(target_dir, '*.docx'))
  sort_list = natsorted(files)
  for filepath in sort_list:
    print (filepath)
    

filepath ()

