import win32api
import sys
import os
import time
from natsort import natsorted

import tkinter as tk
import tkinter.filedialog

def auto_print(path):
  win32api.shellExecute(0, "print", path, None, ".", 0)
 print ("Printed:" + path)


def file_check(path):
  if os.path.isdir(path):
    files = natsorted(os.listdir(path))
    print(files)
    for file in files:
      file_check(path + "\\" + file)
  else:
    auto_print(path)
    time.sleep(3)


dir = os.getcwd()
print_path = tkinter.filedialog.askdirectory(initialdir = dir)
file_check(print_path)

root = tk.Tk()
root.withdraw()

