from pywinauto import Application
from pywinauto.keyboard import send_keys
import os
import time

#input_list = os.listdir("C:\\Users\\Brian Ebrahimi\\Desktop\\Input")

app = Application(backend="uia").start("C:\\Program Files\\Analytic Technologies\\32bit\\Uci.exe")
app.UCINETForWindowsVersion.type_keys("^i^c")
#app.UCINETForWindowsVersion.print_control_identifiers()

import_window = Application(backend='uia').connect(title='Import DL Text File')

for input_file in input_list:
    import_window.ImportDLTextFile.print_control_identifiers()
    import_window.ImportDLTextFile.Edit3.set_edit_text(f"C:\\Users\\Brian Ebrahimi\\Desktop\\Input\\{input_file}")
    if "(" in input_file:
            import_window.ImportDLTextFile.Edit2.set_edit_text(f'C:\\Users\\Brian Ebrahimi\\Desktop\\Input\\{input_file.replace("(", "").replace(")", "").replace(".", "")}')
    import_window.ImportDLTextFile.type_keys("%o")
    time.sleep(0.1)

    try:
        notepad = Application(backend='uia').connect(best_match='ucinetlog- Notepad')
        notepad.ucinetlogNotepad.type_keys("%{F4}")
        app.UCINETForWindowsVersion.set_focus()
        app.UCINETForWindowsVersion.type_keys("^i^c")
    except pywinauto.findbestmatch.MatchError:
        error_window = Application(backend='uia').connect(best_match='UcinetForWindows')
        error_window.UcinetForWindows.set_focus()
        error_window.UcinetForWindows.type_keys('{ENTER}')
        import_window.ImportDLTextFile.set_focus()
        import_window.ImportDLTextFile.type_keys('{BACKSPACE 60}')
