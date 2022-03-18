# Import python libraries
from docx2pdf import convert
from os import listdir
from os.path import isfile, join

# Define pathname of folder that has Word docs
word_folder_pathname = r'C:\path\to\word_docs'
        
# 1) Get Word and PDF files in folder
Word_Files = [f.split('.docx')[0] for f in listdir(word_folder_pathname) if isfile(join(word_folder_pathname, f)) and '.docx' in f and '~' not in f]
PDF_Files = [f.split('.pdf')[0] for f in listdir(word_folder_pathname) if isfile(join(word_folder_pathname, f)) and '.pdf' in f and '~' not in f]

# 2) Isolate Word files that don't have associated PDF
Files_To_Convert = list(set(Word_Files).difference(set(PDF_Files)))
Files_To_Convert.sort()
    
# 3) Convert Word docs to pdf, if this is needed
if len(Files_To_Convert)==0:
     # Do nothing
     print('- No Word files exist or all Word files have an associated PDF file. Exiting script.')
else:
     # Convert to PDF
     print('- ' + str(len(Files_To_Convert)) + ' Word files do not have associated PDF file. Converting them...')
     for file in Files_To_Convert:
         word_file_name = join(word_folder_pathname, file)
         pdf_file_name = join(word_folder_pathname, file)
         convert(word_file_name+'.docx', pdf_file_name+'.pdf')
