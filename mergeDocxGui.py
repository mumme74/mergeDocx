#!/usr/bin/env python3

# Copyright (c) 2023 Fredrik Johansson3 vaxjo se
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

from tkinter import *
from tkinter.filedialog import askdirectory
from mergeDocx import *

# must be global to enable it in open_dir event

class MainForm:
  def clear_log(self):
    self.result_txt.remove("1.0", END)

  def log_fn(self, logStr):
    self.result_txt['state'] = 'normal'
    self.result_txt.insert(END, logStr)
    self.result_txt['state'] = 'disabled'

  def merge_files(self):
    print('merge_files called\n')
    files = list_files(self.dir_path.get())
    if len(files):
      doc = combine_word_documents(files)
      save_file(self.dir_path.get(), doc)
    else:
      self.log_fn('Hittade inga filer i mappen\n')


  def open_dir(self):
    print('open_dir called')
    dirpath = askdirectory()
    if not dirpath:
      return
    if validate_dir(dirpath):
      self.dir_path.set(dirpath)
      self.dir_run_btn['state'] = 'normal'
      print_sep()
      files = list_files(dirpath)
      if len(files):
        self.log_fn('Hittade:\n' + '\n'.join(files))
      else:
        self.log_fn('***Hittade inga docx filer i {}\n'.format(dirpath))
    else:
      self.dir_path.set('Ogiltig sökväg')
      self.dir_run_btn['state'] = 'disabled'


  def __init__(self):
    setPrintFn(self.log_fn)

    self.window = Tk()
    self.window.title('Merge *.docx')
    self.window.rowconfigure(0, minsize=50, weight=1)

    self.dir_path = StringVar(self.window, value="Sökväg till din mapp med docx filer")

    self.dir_frm = Frame(self.window, relief=RAISED)
    self.dir_frm.grid(row=0, column=0, sticky=EW, padx=5, pady=5)
    self.dir_frm.columnconfigure(0, minsize=400)
    self.dir_frm.columnconfigure(1, minsize=50)
    self.dir_frm.columnconfigure(2, minsize=50)

    self.dir_lbl = Label(self.dir_frm, justify=LEFT, text='Mapp med docx filer du vill slå ihop')
    self.dir_lbl.grid(row=0, column=0, sticky=W)
    self.dir_ntr = Entry(self.dir_frm, justify=LEFT, textvariable=self.dir_path, state="readonly", width=50)
    self.dir_ntr.grid(row=1, column=0, sticky=W)
    self.dir_open_btn = Button(self.dir_frm, text='Välj mapp', command=self.open_dir)
    self.dir_open_btn.grid(row=1, column=1)
    self.dir_run_btn = Button(self.dir_frm, text='Slå ihop filer', command=self.merge_files, state="disabled")
    self.dir_run_btn.grid(row=1, column=2)

    self.result_txt = Text(self.window, state="disabled")
    self.result_txt.grid(row=1, column=0)

    self.window.mainloop()


if __name__ == '__main__':
  MainForm()
