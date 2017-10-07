#!/usr/bin/python3
import random
import xlsxwriter
from inspect import getargvalues, currentframe
from os import getcwd
import xlrd

from settings import FILE, GROUPS, MAX_PER_GROUP, NAME_COLUMN, FIRST_CHOICE_COLUMN, SECOND_CHOICE_COLUMN, N_CHOICES, FIRST_DATA_ROW


class File:
  """
  Read a file and return list of entries.
  Entries are tuples containing: (name, first_preference, second_preference).
  """
  def __init__(self, filename, name_column=0, first_choice_column=1, second_choice_column=2, first_data_row=0, random=True, sheet=None):
    """
    Arguments:
      filename (str): The full path to file
      first_data_row (int): The number of the first row that contains user data.
    """

    # Instance attributes
    self.filename = filename
    self.name_column = name_column
    self.first_choice_column = first_choice_column
    self.second_choice_column = second_choice_column
    self.first_data_row = first_data_row > 0 and first_data_row -1 or 0
    self.random = random
    self.sheet = sheet
    self.users = ["File has not (yet) been successfully read."]

  def read_auto(self):
    """
    Read file, guess filetype. Return user data.
    data format: [(name, first_preference, second_preference)]
    """
    ftype = self.filename.split('.')[-1]
    # xlsx and xls are interchangeable for this use
    if ftype == 'xlsx': ftype = 'xls'
    # Check validity of file type.
    if ftype in ('xls', 'ods', 'csv'):
      # Get appropriate method for file type.
      method = getattr(self, 'read_' + ftype)
    else:
      raise TypeError('Cannot read file. csv, xls, xlsx and ods are accepted.')
    # Read file
    self.users = method()
    return self.users

  def read_csv(self):
    """
    Read csv file. Return user data.
    data format: [(name, first_preference, second_preference)]
    """
    # Get rows.
    with open(self.filename) as f:
      rows = []
      for line in f:
        rows.append(line.strip().split(','))

    # Get user data from rows.
    self.users = self._rows_to_users(rows[self.first_data_row:])
    return self.users


  def read_xls(self):
    """
    Read xls or xlsx file. Return user data.
    data format: [(name, first_preference, second_preference)]
    """
    # Get rows.
    book = xlrd.open_workbook(self.filename)
    if not self.sheet:
      sheet = book.sheet_by_index(0)
    else:
      sheet = book.sheet_by_name(self.sheet)

    i = self.first_data_row
    rows = []
    while True:
      try:
        rows.append(sheet.row_values(i))
      except IndexError:
        break
      else:
          i += 1

    # Get user data from rows.
    self.users = self._rows_to_users(rows)
    return self.users

  def _rows_to_users(self, rows):
    """Take rows, return user data.
    data format: [(name, first_preference, second_preference)]
    """
    if self.random:
      random.shuffle(rows)
    return [(row[self.name_column], row[self.first_choice_column], row[self.second_choice_column]) for row in rows]


class Group:

  def __init__(self, users, n_choices, max_per_group):
    """
    Take user data, and allow to assign to groups according to preference
    if possible, at random if not.

    Arguments:
      users (tuple): The users to process (name, first_preference, second_preference)
      n_choices (int): The number of alternatives people can choose between.
      max_per_group (int): The maximum allowed people per group.
    """

    self.n_choices = n_choices
    self.max_per_group = max_per_group
    self.users = users

    self.groups = {}

  def assign(self):
    """ Assign users to groups
    Return groups (dict). """

    # Clear groups
    self.groups = {}
    for i in range(1, self.n_choices+1):
      self.groups[i] = []
    users = self.users[:]

    # Fill groups
    while users:
      u = users.pop()
      name = u[0]
      #First choice
      first_choice = int(u[1])
      if len(self.groups[first_choice]) < self.max_per_group:
        self.groups[first_choice].append(name)
        continue
      # Second choice
      second_choice = int(u[2])
      if len(self.groups[second_choice]) < MAX_PER_GROUP:
        self.groups[second_choice].append(name)
        continue

      # If first and second choices are full, assign at random
      while u:
        i = random.randrange(1,self.max_per_group+1)
        if len(self.groups[i]) < self.max_per_group:
          self.groups[i].append(name)
          u = []

    return self.groups

  def write_to_file(self, groups=None):
    """ Write groups to xlsx file.
    Return message (str)."""

    if not groups:
      groups = self.groups
    if not groups:
      return "Cannot write to file, users are not assigned to groups (correctly)."

    # Create file
    workbook = xlsxwriter.Workbook('results.xlsx')
    worksheet = workbook.add_worksheet()
    # Fill file
    row = 0
    col = 0

    for group, users in groups.items():
      worksheet.write(row, col, 'Groep ' + str(group))
      for user in users:
        row += 1
        worksheet.write(row,col,user)
      row = 0
      col += 1

    # Release data
    workbook.close()
    return "File created succesfully at %s/results.xlsx" % getcwd()


if __name__ == '__main__':
  f = File(FILE, NAME_COLUMN, FIRST_CHOICE_COLUMN, SECOND_CHOICE_COLUMN, FIRST_DATA_ROW)
  data = f.read_auto()
  groups = Group(data, N_CHOICES, MAX_PER_GROUP)
  groups.assign()
  print(groups.write_to_file())
