import PySimpleGUI as sg
import json



def win2():
    sg.set_options(font=('Arial Bold', 16))
    layout = [
    [sg.Text('Settings', justification='left')],
    [sg.Text('User name', size=(10, 1), expand_x=True),
    sg.Input(key='-USER-')],
    [sg.Text('email ID', size=(10, 1), expand_x=True),
    sg.Input(key='-ID-')],
    [sg.Text('Role', size=(10, 1), expand_x=True),
    sg.Input(key='-ROLE-')],
    [sg.Button("LOAD"), sg.Button('SAVE'), sg.Button('Exit')]
    ]
    window = sg.Window('User Settings Demo', layout, size=(715, 200))
    # Event Loop
    while True:
        event, values = window.read()
        if event in (sg.WIN_CLOSED, 'Exit'):
            break
        if event == 'LOAD':
            f = open("settings2.txt", 'r')
            settings = json.load(f)
            window['-USER-'].update(value=settings['-USER-'])
            window['-ID-'].update(value=settings['-ID-'])
            window['-ROLE-'].update(value=settings['-ROLE-'])
        if event == 'SAVE':
            settings = {'-USER-': values['-USER-'],
            '-ID-': values['-ID-'],
            '-ROLE-': values['-ROLE-']}
            f = open("settings2.txt", 'w')
            json.dump(settings, f)
            f.close()
    window.close()

def open_window():
    layout = [[sg.Text("New Window", key="new")]]
    window = sg.Window("Second Window", layout, modal=True)
    choice = None
    while True:
        event, values = window.read()
        if event == "Exit" or event == sg.WIN_CLOSED:
            break
        
    window.close()


def main():
    layout = [[sg.Button("Open Window", key="open")]]
    window = sg.Window("Main Window", layout)
    while True:
        event, values = window.read()
        if event == "Exit" or event == sg.WIN_CLOSED:
            break
        if event == "open":
            win2()
        
    window.close()


if __name__ == "__main__":
    main()