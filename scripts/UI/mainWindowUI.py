import tkinter as tk
import webbrowser

import scripts.gvar as gvar
import scripts.UI.browseFilesWindow as browseFilesWindow
import scripts.configuration.configuration as configuration
import scripts.configuration.DDMMFormat as DDMMFormat
import scripts.utilities.exit as exit
import scripts.executeConversionToXlsx as executeConversionToXlsx

def mainWindowUI() -> None:
    """main window UI"""
    gvar.window = tk.Tk()
    gvar.window.title('KS Project')
    gvar.window.geometry("700x310")
    gvar.window.config(background = "turquoise2")

    gvar.label_file_explorer = tk.Label(gvar.window, 
							text = "KS-ChatAnalyzer",
							width = 44, height = 2, 
							fg = "black",
                            background="pale green",
                            font=("Arial", 20)
        )

    button_explore = tk.Button(gvar.window, 
                        text = "Browse Files",
                        command = browseFilesWindow.browseFiles,
                        width = 40, height = 2
                        )

    button_run = tk.Button(gvar.window, 
						text = "Run program",
						command = executeConversionToXlsx.runProgram,
                        width = 40, height = 2)
    
    button_config = tk.Button(gvar.window, 
						text = "Configuration",
						command = configuration.configurationMenu,
                        width = 40, height = 2) 
    
    button_exit = tk.Button(gvar.window, 
					text = "Exit",
					command = exit.exit,
                    width = 40, height = 2) 

    gvar.label_file_explorer.grid(column = 1, row = 1, columnspan=2)

    button_explore.grid(column = 1, row = 2, columnspan=2)

    button_run.grid(column = 1, row = 4, columnspan=2)
    
    button_config.grid(column = 1, row = 5, columnspan=2)

    button_exit.grid(column = 1,row = 6, columnspan=2)

    dd_mmFormat = tk.BooleanVar(value=gvar.config["dd_mmFormat"])
    DDMM_Button = tk.Radiobutton(gvar.window, text="DD_MM_YY Format", variable=dd_mmFormat,
                                indicatoron=False, value=True, width=19, height = 2, command=DDMMFormat.setDDMM)
    MMDD_Button = tk.Radiobutton(gvar.window, text="MM_DD_YY Format", variable=dd_mmFormat,
                                indicatoron=False, value=False, width=19, height = 2, command=DDMMFormat.setMMDD)
    
    DDMM_Button.grid(column = 1,row = 3, sticky="e")
    MMDD_Button.grid(column = 2,row = 3, sticky="w")
    
    #footer
    footerLabel = tk.Label(gvar.window, 
                            text = "Author: github.com/roccat1",
                            width = 87, height = 2, 
                            fg = "black",
                            background="pale green",
                            font=("Arial", 10)
        )
    
    footerLabel.bind("<Button-1>", lambda e:webbrowser.open_new_tab("https://github.com/roccat1"))
    
    footerLabel.grid(column = 1, row = 7, columnspan=2, sticky="w")
    
    # Let the window wait for any events
    gvar.window.mainloop()