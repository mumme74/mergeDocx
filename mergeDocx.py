#! /usr/bin/env python3

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

import sys, os
from docx import Document

my_print = print
def setPrintFn(fn):
  global my_print
  my_print = fn

def combine_word_documents(input_files):
  """
  :param input_files: an iterable with full paths to docs
  :return: a Document object with the merged files
  """
  print_sep()

  for filnr, file in enumerate(input_files):
    my_print("{0} {1}\n".format(filnr, file))

    if filnr == 0:
      merged_document = Document(file)
      merged_document.add_page_break()

    else:
      sub_doc = Document(file)

      # Don't add a page break if you've reached the last file.
      if filnr < len(input_files)-1:
        sub_doc.add_page_break()

      for element in sub_doc.element.body:
        merged_document.element.body.append(element)

  return merged_document

def list_files(dir):
  files = [ os.path.abspath(os.path.join(dir, f))
              for f in os.listdir(dir)
                if f.endswith('.docx') and not f.startswith('merged')]
  return files

def save_file(dir, doc):
  nr = 0
  path = lambda : os.path.join(dir, 'merged{0}.docx'.format(nr))

  while os.path.exists(os.path.join(dir, path())):
    nr += 1
  doc.save(path())
  print_sep()
  my_print("Har skrivit till filen 'merged{0}.docx' i samma mapp\n".format(nr))

def print_sep():
  my_print("\n--------------------------------------\n")

def validate_dir(dir):
  if not os.path.exists(dir):
    my_print("Sökvägen '{0}' finns inte\n".format(dir))
  elif not os.path.isdir(dir):
    my_print("Sökvägen '{0}' är inte en mapp\n".format(dir))
  else:
    return True
  return False

if __name__ == '__main__':
    if len(sys.argv) != 2:
      my_print('{0} path/to/folder/\n'.format(sys.argv[0]) +
            ' Du måste ge en sökväg till mappen med docx filerna du vill merga.\n')
    elif validate_dir(sys.argv[1]):
      files = list_files(sys.argv[1])
      my_print(files)
      doc = combine_word_documents(files)
      save_file(sys.argv[1], doc)
    else:
      sys.exit(1)

