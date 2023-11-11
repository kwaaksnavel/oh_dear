from win32com.client import gencache, pythoncom
import os, sys
import re
from pathlib import Path
import tkinter as tk
from tkinter import filedialog
import pandas as pd
import shutil
import random


def clippy():
    print('                        _.-;:q=._                                    ')
    print('                      .\' j=""^k;:\.                                  ')
    print('                     ; .F       ";`Y                                 ')
    print('                    ,;.J_        ;\'j                                ')
    print('                  ,-;"^7F       : .F           _________________    ')
    print('                 ,-\'-_<.        ;gj. _.,---""''               .\'   ')
    print('                ;  _,._`\.     : `T"5,                       ;      ')
    print('                : `?8w7 `J  ,-\'" -^q. `                     ;    ')
    print('                 \;._ _,=\' ;   n58L Y.                     .\'    ')
    print('                   F;";  .\' k_ `^\'  j\'                     ; ')
    print('                   J;:: ;     "y:-=\'                      ;  ')
    print('                    L;;==      |:;   jT\                  ;  ')
    print('                    L;:;J      J:L  7:;\'       _         ;   ')
    print('                    I;|:.L     |:k J:.', '       .     ;   ')
    print('                    |;J:.|     ;.I F.:      .           :    ')
    print('                   ;J;:L::     |.| |.J  , \'   `    ;    ;    ')
    print('                 .\' J:`J.`.    :.J |. L .    ;         ; ')
    print('                ;    L :k:`._ ,\',j J; |  ` ,        ; ;  ')
    print('              .\'     I :`=.:."_".\'  L J             `.\'  ')
    print('            .\'       |.:  `"-=-\'    |.J              ;   ')
    print('        _.-\'         `: :           ;:;           _ ;    ')
    print('    _.-\'"             J: :         /.;\'       ;    ; ')
    print('  =\'_                  k;.\.    _.;:Y\'     ,     .\'  ')
    print('     `"---..__          `Y;."-=\';:=\'     ,      .\'   ')
    print('              `""--..__   `"==="\'    -        .\' ')
    print('                       ``""---...__        .-\'   ')
    print('                                   ``""---\'  ')


class Woord:
    def __init__(self, visible=0, scr_upd=0, toc_upd=False, toc_ref=False):
        pythoncom.CoInitialize()
        self.app = gencache.EnsureDispatch('Word.Application')
        self.app.Visible = visible
        self.app.DisplayAlerts = 0
        self.app.ScreenUpdating = scr_upd
        self.toc_upd = toc_upd
        self.toc_ref = toc_ref

    def body_highlight(self, fallback=False):
        # Create word list
        word_l = keywords
        # Pruning non-matching words from word list if everything went well
        if not fallback:
            word_l = [[k for k, v in dict.items() if v > 0] for dict in self.word_d]
            print(f'Found {sum([sum(d.values()) for d in self.word_d])} instances of {sum([len(lijst) for lijst in word_l])} keyword(s)')

        # Highlight the terms in the list
        print('>> Highlighting Words in Text')
        self.palette = [3, 4, 5, 7, 6, 2, 11, 12, 14, 13, 0]  # highlight colours, cant rgb! only about 10 working colours...
        self.app.Selection.Find.Replacement.Highlight = True  # turn on highlighting
        self.app.Selection.GoTo(What=0, Which=1)

        for k in range(len(word_l)):
            for word in word_l[k]:
                self.app.Options.DefaultHighlightColorIndex = self.palette[k]  # picking colour
                self.app.Selection.Find.Text = word  # finding word matches in the document
                self.app.Selection.Find.Replacement.Text = word
                self.app.Selection.Find.Execute(Replace=2, MatchWholeWord=False, MatchCase=False) # changed to False for Tom

    def highlight(self, docpath, keywords, tmp_dir):
        self.opendoc = os.path.basename(docpath)
        docu = self.app.Documents.Open(FileName=docpath)
        print('>> Opening D80 Assessment Report')
        self.body_highlight(fallback=True)

        return docu

    def save_docx(self, root, dear, ohdear):

        doc_name = re.split('(^.*\\\)', dear)[2]
        #name = re.split('\.', doc_name)[0]
        ext = re.split('\.', doc_name)[1]

        try:
            if ext == 'docx':
                if self.toc_upd:
                    ohdear.SaveAs(f'{root}\\ohdear_ToC_upd_{doc_name}', FileFormat=16)
                else:
                    ohdear.SaveAs(f'{root}\\ohdear_{doc_name}', FileFormat=16)
            elif ext == 'doc':
                if self.toc_upd:
                    ohdear.SaveAs(f'{root}\\ohdear_ToC_upd_{doc_name}', FileFormat=0)
                else:
                    ohdear.SaveAs(f'{root}\\ohdear_{doc_name}', FileFormat=0)

        except:
            print('ERROR: Unable to save. Make sure file with same name is not already opened')

        if self.app.Visible is False:
            self.app.ActiveDocument.Close()
            self.app.Quit()

        return doc_name

# ohdear: Overly highlighting day eighty assessment reports
if __name__ == "__main__":
    print('ohdear: Overly Highlighting Day Eighty Assessment Reports')
    print('       __        __')
    print(' ___  / /    ___/ /__ ___ _____')
    print('/ _ \/ _ \  / _  / -_) _ `/ __/')
    print('\___/_//_/  \_,_/\__/\_,_/_/   ')
    print('')

    global DEBUG
    DEBUG = False
    TOC_UPD = False

    # Cant hurt
    import time
    then = time.time()

    # Get the home directory
    home = str(Path.home())
    tmp_dir = home + r'\AppData\Local\Temp\\'

    # Remove gen_py folder to prevent incidental errors
    gen_py = Path(f'{tmp_dir}gen_py')
    if gen_py.exists() and gen_py.is_dir():
        shutil.rmtree(gen_py)

    # Get location of excel sheet
    root = tk.Tk()
    root.withdraw()
    root.focus_force()
    # DEAR = filedialog.askopenfilename(initialdir = rf'{home}\Documents',
    #                                   title = 'Select a Day Eighty Assessment Report',
    #                                   filetypes = (('word documents', '*.docx'),('all files','*.*')))

    print('Select Day Eighty Assessment Report Word File in Popup Window')

    dear = filedialog.askopenfilename(initialdir = rf'{home}\Downloads',
                                      title = 'Select a Day Eighty Assessment Report')

    dear = re.sub('/', '\\\\', dear) # so we can have spaces in document names


    # Load keywords from excel file
    if getattr(sys, 'frozen', False):  # so we can fetch the location where .exe is run
        application_path = os.path.dirname(sys.executable)
        os.chdir(application_path)

    print('>> Collecting Keywords')  # requires openpyxl package installed to work
    root = os.getcwd()
    try:
        sheet = pd.read_excel(f'{root}\\keywords.xlsx', skiprows=1, header=None,
                              na_filter=True, dtype=str)

    except:
        print('ERROR: Keywords not found. Make sure keywords excel file is in the same folder as ohdear')
        os.system('pause')
        sys.exit()

    sheet.fillna('', inplace=True)
    keys = sheet.transpose().values.tolist()

    keywords = [[x for x in lijst if x] for lijst in keys]

    while keywords[-1] == []:
        keywords = keywords[:-1]

    # Open doc behind the scenes and save as txt to perform regex on later
    print('>> Initialising Win32 COM API')
    doc_txt = Woord(visible=0, scr_upd=0, toc_upd=TOC_UPD)
    # doc_txt.save_txt(dear, tmp_dir)

    # Open version that will be highlighted and saved in directory .exe was run from
    if DEBUG:
        doc = Woord(visible=1, scr_upd=1, toc_upd=TOC_UPD)
    else:
        doc = Woord(visible=0, scr_upd=0, toc_upd=TOC_UPD)

    ohdear = doc.highlight(dear, keywords, tmp_dir)

    doc_name = doc.save_docx(root, dear, ohdear)
    print(f'>> Saving as ohdear_{doc_name} in same directory as ohdear')
    # Easter egg
    if (random.random()) > 0.99:
        clippy()
        print('Clippy suggests: "Ctrl+F can also help you look for keywords in the text!"')

    # Calculate how much of a person's time we have wasted
    now = time.time()
    print('ohdear took', round(now - then, 1), 'seconds')

    # To stop the window from closing
    # os.remove(f'{tmp_dir}workaround.txt')  # remove temporary file
    os.system('pause')
