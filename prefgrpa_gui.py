#!/usr/bin/python3
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk

import xlrd

from prefgrpa import Group, File

class App:
  frame = None

  def __init__(self, master):
    self.master = master
    self.frame = None
    self.fname = ''
    self.columns = []
    self.select_file()
    self.sheet = None

  def clear_frame(self):
    """ Clear window frame. """
    self.frame.destroy()
    self.frame = tk.Frame(self.master)
    self.frame.pack()


  def select_file(self):
    """ Select file frame. """
    if self.frame:
      self.clear_frame()

    self.frame = tk.Frame(self.master)
    self.frame.pack()
    r = 0
    tk.Label(self.frame, text="""The input file can be .csv, .xls, .xlsx, .ods
             For csv only comma separated is allowed.""").pack()
    r += 1
    tk.Button(self.frame, text="Choose file to read", command=self._load_file).pack()
    r += 1
    self.file_label = tk.StringVar()
    tk.Label(self.frame, textvariable=self.file_label).pack()
    r += 1
    tk.Label(self.frame, text="Written by Dominic Dingena").pack()

  def _load_file(self):
    """ Load selected file. """
    self.fname = filedialog.askopenfilename()
    if self.fname:
      self.file_label.set(self.fname)

    extension = self.fname.split('.') [-1]
    if extension == 'xlsx':
      extension = 'xls'

    if extension == 'csv':
      self.get_columns_csv()

    if extension == 'xls':
      book = xlrd.open_workbook(self.fname)
      sheets = book.sheet_names()
      self.clear_frame()
      tk.Label(self.frame, text='Select sheet:').pack()
      l = tk.Listbox(self.frame, selectmode=tk.SINGLE)
      for item in sheets:
        l.insert(tk.END, item)
      l.pack()
      tk.Button(self.frame, text="Continue", command=lambda: self.get_columns_xls(l.get(l.curselection()))).pack()

  def get_columns_csv(self):
    """ Get the letters for the available columns. """
    with open(self.fname) as f:
      # Get number of columns
      length = len(f.readline().strip().split(','))
    # Create letters for column
    columns = [chr(i+65) for i in range(length)]
    self.set_options(columns)

  def get_columns_xls(self, sheet):
    """ Get the letters for the available columns in xls(x) file. """
    self.sheet = sheet
    i = 0
    book = xlrd.open_workbook(self.fname)
    sheet = book.sheet_by_name(sheet)
    col_values = [1]

    while col_values:
      try:
        col_values = sheet.col_values(i)
      except IndexError:
        col_values = []
      else:
        i+=1
    columns = [chr(x+65) for x in range(i)]
    self.set_options(columns)

  def set_options(self, columns):
    """ Set options frame. """
    self.clear_frame()
    r = 0
    tk.Label(self.frame, text='Number of choices:').grid(row=r, column=0)
    self.n_choices = tk.Entry(self.frame, width=3)
    self.n_choices.grid(row=r, column=1)
    r += 1
    tk.Label(self.frame, text='Max per group:').grid(row=r, column=0)
    self.max_per_group = tk.Entry(self.frame, width=3)
    self.max_per_group.grid(row=r, column=1)
    r += 1
    tk.Label(self.frame, text='Number of the first row containing user data').grid(row=r, column=0)
    self.first_data_row = tk.Entry(self.frame, width=3)
    self.first_data_row.grid(row=r, column=1)
    r += 1
    tk.Label(self.frame, text='Column containing persons names').grid(row=r, column=0)
    self.name_column = ttk.Combobox(self.frame)
    self.name_column['values'] = columns
    self.name_column.grid(row=r, column=1)
    r += 1
    tk.Label(self.frame, text='Column containing first preference').grid(row=r, column=0)
    self.first_choice = ttk.Combobox(self.frame)
    self.first_choice['values'] = columns
    self.first_choice.grid(row=r, column=1)
    r += 1
    tk.Label(self.frame, text='Column containing second preference').grid(row=r, column=0)
    self.second_choice = ttk.Combobox(self.frame)
    self.second_choice['values'] = columns
    self.second_choice.grid(row=r, column=1)
    r += 1
    tk.Button(self.frame, text="Back", command=self.select_file).grid(row=r, column=0)
    tk.Button(self.frame, text="Generate file", command=self.generate_choices).grid(row=r, column=1)
    r += 1
    tk.Label(self.frame, text="Written by Dominic Dingena").grid(row=r, columnspan=2)
    pass

  def generate_choices(self):
    """ Assign users to groups based on the file and options specified. """
    # Letters to number
    try:
      name_column = ord(self.name_column.get().lower()) - 97
      first_choice = ord(self.first_choice.get().lower()) - 97
      second_choice = ord(self.second_choice.get().lower()) - 97
      first_data_row = int(self.first_data_row.get())
      filename = self.fname
      n_choices = int(self.n_choices.get())
      max_per_group = int(self.max_per_group.get())
    except Exception:
      messagebox.showerror('Error', """The information you entered is not valid.
Please follow the instructions.""")
    else:
      filename = self.fname
      sheet = self.sheet
      f = File(filename, name_column, first_choice, second_choice, first_data_row, sheet)
      users = f.read_auto()
      g = Group(users, n_choices, max_per_group)
      g.assign()
      message = g.write_to_file()
      messagebox.showinfo('Succes.', message)

def main():
  """ Initialize tkinter. """
  root = tk.Tk()
  root.wm_title("Automated group selection")
  app = App(root)
  root.mainloop()
  #root.destroy()

if __name__ == '__main__':
  main()
