from pywinauto import Application
from pywinauto.keyboard import send_keys
import os
import time
from pywinauto.mouse import click

input_list_initial = os.listdir("C:\\Users\\Brian Ebrahimi\\Desktop\\ucinetautomator\\output")
input_list_final = [file for file in input_list_initial if file.endswith(".##h")]

app = Application(backend="uia").start("C:\\Program Files (x86)\\Analytic Technologies\\Uci.exe")

app.UCINETForWindowsVersion.type_keys("%n")
app.UCINETForWindowsVersion.type_keys("w")
app.UCINETForWindowsVersion.type_keys("c")
#app.UCINETForWindowsVersion.print_control_identifiers()

import_window = Application(backend='uia').connect(title='Clustering Coefficient')
#import_window.ClusteringCoefficient.print_control_identifiers()

for input_file in input_list_final:
    import_window.ClusteringCoefficient.print_control_identifiers()
    import_window.ClusteringCoefficient.Edit1.set_edit_text(f"C:\\Users\\Brian Ebrahimi\\Desktop\\output\\{input_file}")
    import_window.ClusteringCoefficient.Edit2.set_edit_text(f"C:\\Users\\Brian Ebrahimi\\Desktop\\ucinetautomator\\output\\{input_file}")
    import_window.ClusteringCoefficient.OK.click()
    time.sleep(0.1)

    notepad = Application(backend='uia').connect(best_match='ucinetlog- Notepad')
    notepad.ucinetlogNotepad.type_keys("%{F4}")

    app.UCINETForWindowsVersion.set_focus()
    app.UCINETForWindowsVersion.type_keys("%n")
    app.UCINETForWindowsVersion.type_keys("w")
    app.UCINETForWindowsVersion.type_keys("c")