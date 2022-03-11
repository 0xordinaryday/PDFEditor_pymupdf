#!/usr/bin/env python
"""

"""
import PyPDF2
import PySimpleGUI as sg
import sys
import fitz

if sys.platform == "win32":
    import ctypes

    ctypes.windll.shcore.SetProcessDpiAwareness(2)


def get_page(doc, pno, max_size=None, dlist=None):
    dlist_tab = []
    """Return a tkinter.PhotoImage or a PNG image for a document page number.
    :arg int pno: 0-based page number
    :arg zoom: top-left of old clip rect, and one of -1, 0, +1 for dim. x or y
               to indicate the arrow key pressed
    :arg max_size: (width, height) of available image area
    :arg bool first: if True, we cannot use tkinter
    """
    dlist_tab = [None] * 2  # get display list of page number
    if not dlist:  # create if not yet there
        dlist_tab[pno] = doc[pno].get_displaylist()
        dlist = dlist_tab[pno]
    r = dlist.rect  # the page rectangle
    clip = r
    # ensure image fits screen:
    # exploit, but do not exceed width or height
    zoom_0 = 0.5

    mat_0 = fitz.Matrix(zoom_0, zoom_0)

    pix = dlist.get_pixmap(matrix=mat_0, alpha=False)
    img = pix.tobytes("ppm")  # make PPM image from pixmap for tkinter
    return img, clip.tl  # return image, clip position


def process_file(fname):
    doc = fitz.open(fname)
    page_count = len(doc)
    # print(page_count)
    dlist_tab = [None] * page_count

    SOME_VAL = 50
    max_size = (SOME_VAL, SOME_VAL)

    cur_page = 0
    mydata, clip_pos = get_page(doc, cur_page, max_size=max_size)
    return mydata, clip_pos


def make_preview():
    preview_column = sg.Column([
        [sg.Frame('Preview:', [[sg.Image(key='-PREVIEW-', size=(300, 420))],
                               [sg.ReadFormButton("Prev"), sg.ReadFormButton("Next")],
                               [sg.Text("Page:"), sg.InputText(
                                   "Page", size=(5, 1), do_not_clear=True, key="-PageNumber-"),
                                # str(cur_page + 1), size=(5, 1), do_not_clear=True, key="-PageNumber-"
                                sg.Text("Pages")]], element_justification='c', title_location='n')], ], pad=(0, 0),
        element_justification='c')
    return preview_column


def make_window(theme):
    sg.theme(theme)
    menu_def = [['&Application', ['&Exit']],
                ['&Help', ['&About']]]
    right_click_menu_def = [[], ['Edit Me', 'Versions', 'Nothing', 'More Nothing', 'Exit']]
    graph_right_click_menu_def = [[], ['Erase', 'Draw Line', 'Draw', ['Circle', 'Rectangle', 'Image'], 'Exit']]

    # Table Data
    data = [["John", 10], ["Jen", 5]]
    headings = ["Name", "Score"]

    extract_layout_column1 = sg.Column([
        [sg.Frame('Categories:', [[sg.Radio('Websites', 'radio1', default=True, key='-WEBSITES-', size=(10, 1)),
                                   sg.Radio('Software', 'radio1', key='-SOFTWARE-', size=(10, 1))]], )],
        # Information sg.Frame
        [sg.Frame('Information:', [
            [sg.Text('Anything that requires user-input is in this tab!')],
            [sg.Input(key='-INPUT-')],
            [sg.Checkbox('Checkbox', default=True, k='-CB-')],
            [sg.Radio('Radio1', "RadioDemo", default=True, size=(10, 1), k='-R1-'),
             sg.Radio('Radio2', "RadioDemo", default=True, size=(10, 1), k='-R2-')],
            [sg.Combo(values=('Combo 1', 'Combo 2', 'Combo 3'), default_value='Combo 1', readonly=True, k='-COMBO-'),
             sg.OptionMenu(values=('Option 1', 'Option 2', 'Option 3'), k='-OPTION MENU-'), ],
            [sg.Spin([i for i in range(1, 11)], initial_value=10, k='-SPIN-'), sg.Text('Spin')],
            [sg.Multiline(
                'Demo of a Multi-Line Text Element!\nLine 2\nLine 3\nLine 4\nLine 5\nLine 6\nLine 7\nYou get the point.',
                size=(45, 5), expand_x=False, expand_y=True, k='-MLINE-')],
            [sg.Button('Open File'), sg.Button('Popup')]])], ], pad=(0, 0), element_justification='c')

    extract_layout = [[extract_layout_column1, make_preview()]]

    aesthetic_layout_column1 = sg.Column([
        [sg.Frame('Categories:', [[sg.T('Anything that you would use for asthetics is in this tab!')],
                       [sg.ProgressBar(100, orientation='h', size=(20, 20), key='-PROGRESS BAR-'),
                        sg.Button('Test Progress bar')]], pad=(0, 0), element_justification='c')]])

    asthetic_layout = [[aesthetic_layout_column1, make_preview()]]

    logging_layout = [[sg.Text("Anything printed will display here!")],
                      [sg.Multiline(size=(60, 15), font='Courier 8', expand_x=False, expand_y=True, write_only=True,
                                    reroute_stdout=True, reroute_stderr=True, echo_stdout_stderr=True, autoscroll=True,
                                    auto_refresh=True)]]

    graphing_layout = [[sg.Text("Anything you would use to graph will display here!")],
                       [sg.Graph((200, 200), (0, 0), (200, 200), background_color="black", key='-GRAPH-',
                                 enable_events=True,
                                 right_click_menu=graph_right_click_menu_def)],
                       [sg.T('Click anywhere on graph to draw a circle')],
                       [sg.Table(values=data, headings=headings, max_col_width=25,
                                 background_color='black',
                                 auto_size_columns=True,
                                 display_row_numbers=True,
                                 justification='right',
                                 num_rows=2,
                                 alternating_row_color='black',
                                 key='-TABLE-',
                                 row_height=25)]]

    popup_layout = [[sg.Text("Popup Testing")],
                    [sg.Button("Open Folder")],
                    [sg.Button("Open File")]]

    theme_layout = [[sg.Text("See how elements look under different themes by choosing a different theme here!")],
                    [sg.Listbox(values=sg.theme_list(),
                                size=(20, 12),
                                key='-THEME LISTBOX-',
                                enable_events=True)],
                    [sg.Button("Set Theme")]]

    layout = [[sg.MenubarCustom(menu_def, background_color='white', text_color='black', key='-MENU-',
                                font=('Roboto', 10), tearoff=False)]]
    # [sg.Menu(menu_def, background_color=('white'), text_color=('black'), font=('Roboto', 6))],
    layout += [[sg.TabGroup([[sg.Tab('Extract', extract_layout),
                              sg.Tab('Merge', asthetic_layout),
                              sg.Tab('Replace', graphing_layout),
                              sg.Tab('Delete', popup_layout),
                              sg.Tab('Explode', theme_layout),
                              sg.Tab('Output', logging_layout)]], key='-TAB GROUP-', expand_x=True, expand_y=True),

                ]]
    layout[-1].append(sg.Sizegrip())
    window = sg.Window('PDF Editor', layout, right_click_menu=right_click_menu_def,
                       right_click_menu_tearoff=True, grab_anywhere=True, resizable=True, margins=(0, 0),
                       use_custom_titlebar=True, finalize=True, keep_on_top=True,
                       # scaling=2.0,
                       )
    window.set_min_size(size=(700, 400))
    return window


def main():
    window = make_window(sg.theme())

    # This is an Event Loop
    while True:
        event, values = window.read(timeout=100)
        # keep an animation running so show things are happening
        if event not in (sg.TIMEOUT_EVENT, sg.WIN_CLOSED):
            print('============ Event = ', event, ' ==============')
            print('-------- Values Dictionary (key=value) --------')
            for key in values:
                print(key, ' = ', values[key])
        if event in (None, 'Exit'):
            print("[LOG] Clicked Exit!")
            break
        elif event == 'About':
            print("[LOG] Clicked About!")
            sg.popup('This is a thing',
                     "Don't ask", keep_on_top=True)
        elif event == 'Popup':
            print("[LOG] Clicked Popup Button!")
            sg.popup("You pressed a button!", keep_on_top=True)
            print("[LOG] Dismissing Popup!")
        elif event == 'Test Progress bar':
            print("[LOG] Clicked Test Progress Bar!")
            progress_bar = window['-PROGRESS BAR-']
            for i in range(100):
                print("[LOG] Updating progress bar by 1 step (" + str(i) + ")")
                progress_bar.update(current_count=i + 1)
            print("[LOG] Progress bar complete!")
        elif event == "-GRAPH-":
            graph = window['-GRAPH-']  # type: sg.Graph
            graph.draw_circle(values['-GRAPH-'], fill_color='yellow', radius=20)
            print("[LOG] Circle drawn at: " + str(values['-GRAPH-']))
        elif event == "Open Folder":
            print("[LOG] Clicked Open Folder!")
            folder_or_file = sg.popup_get_folder('Choose your folder', keep_on_top=True)
            sg.popup("You chose: " + str(folder_or_file), keep_on_top=True)
            print("[LOG] User chose folder: " + str(folder_or_file))
        elif event == "Open File":
            print("[LOG] Clicked Open File!")
            # folder_or_file = sg.popup_get_file('Choose your file', keep_on_top=True)
            fname = sg.popup_get_file("Select file and filetype to open:", title="PyMuPDF Document Browser",
                                      file_types=(("PDF Files", "*.pdf"),), keep_on_top=True)
            mydata, clip_pos = process_file(fname)
            print(type(mydata))
            window.Element('-PREVIEW-').Update(data=mydata)
            sg.popup("You chose: " + str(fname), keep_on_top=True)
            print("[LOG] User chose file: " + str(fname))
        elif event == "Set Theme":
            print("[LOG] Clicked Set Theme!")
            theme_chosen = values['-THEME LISTBOX-'][0]
            print("[LOG] User Chose Theme: " + str(theme_chosen))
            window.close()
            window = make_window(theme_chosen)
        elif event == 'Edit Me':
            sg.execute_editor(__file__)
        elif event == 'Versions':
            sg.popup(sg.get_versions(), keep_on_top=True)

    window.close()
    exit(0)


if __name__ == '__main__':
    sg.theme('DefaultNoMoreNagging')  # sg.theme('DarkGrey7')
    main()
