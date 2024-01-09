import os
from tkinter import filedialog

import scripts.gvar as gvar

def browseFiles() -> None:
    """Opens a file explorer to select a file which path is saved in the global variable filename"""

    gvar.filename = filedialog.askopenfilename(initialdir = os.getcwd(),
                                            title = "Select a File",
                                            filetypes = (("Text files",
                                            "*.txt*"),
                                            ("all files",
                                            "*.*")))
	
	# Change label contents
    gvar.label_file_explorer.configure(text="File Opened: "+os.path.basename(gvar.filename))
