#!/usr/bin/env python
"""

"""
import PySimpleGUI as sg
import sys
import fitz
import os

display_list_table = []
merge_file_list = []

if sys.platform == "win32":
    import ctypes
    ctypes.windll.shcore.SetProcessDpiAwareness(2)


def get_page(pdf_document, pagenumber):
    """Return a tkinter.PhotoImage or a PNG image for a document page number.
    :param dlist:
    :arg int pno: 0-based page number
    :arg zoom: top-left of old clip rect, and one of -1, 0, +1 for dim. x or y
               to indicate the arrow key pressed
    :arg max_size: (width, height) of available image area
    :arg bool first: if True, we cannot use tkinter
    """
    global display_list_table
    dlist = display_list_table[pagenumber]  # get display list of page number
    if not dlist:  # create if not yet there
        display_list_table[pagenumber] = pdf_document[pagenumber].get_displaylist()
        dlist = display_list_table[pagenumber]
    r = dlist.rect  # the page rectangle
    clip = r

    SIZE_VAL = 495

    if clip.width / clip.height < SIZE_VAL / SIZE_VAL:
        # clip is narrower: zoom to window HEIGHT
        zoom = SIZE_VAL / clip.height
    else:  # clip is broader: zoom to window WIDTH
        zoom = SIZE_VAL / clip.width
    mat = fitz.Matrix(zoom, zoom)

    pix = dlist.get_pixmap(matrix=mat, clip=clip, alpha=False)
    current_image = pix.tobytes("ppm")  # make PPM image from pixmap for tkinter
    return current_image, pagenumber  # return image, don't care about clip position


def is_enter(btn):
    return btn.startswith("Return:") or btn == chr(13)


def is_quit(btn):
    return btn == chr(27) or btn.startswith("Escape:")


def is_next(btn):
    return btn.startswith("Next") or btn == "MouseWheel:Down"


def is_prev(btn):
    return btn.startswith("Prev") or btn == "MouseWheel:Up"


def is_mykeys(btn):
    return any((is_enter(btn), is_next(btn), is_prev(btn)))


def process_file(fname):
    global display_list_table
    pdf_document = fitz.open(fname)
    page_count = len(pdf_document)
    display_list_table = [None] * page_count

    current_page = 0
    current_image, pagenumber = get_page(pdf_document, current_page)
    return pdf_document, current_image, page_count, pagenumber


def make_preview():
    preview_frame = [sg.Frame('Preview:', [[sg.Image(key='-PREVIEW-', size=(495, 495), expand_x=False, expand_y=False)],
                                           [sg.ReadFormButton("Prev"), sg.ReadFormButton("Next")],
                                           [sg.Text("Page:"), sg.Text("1 of", key="-PAGENUMBER-"),
                                            # str(cur_page + 1), size=(5, 1), do_not_clear=True, key="-PageNumber-"
                                            sg.Text("Pages", key="-TotalPages-")]], element_justification='c',
                              title_location='n',
                              key="-PREVIEWFRAME-", visible=False)]
    return preview_frame


def make_main_frame():
    main_frame = [sg.Frame('Functions:', [[
        sg.Button('Extract', key='-EXTRACT-'),
        sg.Button('Delete', key='-DELETE-'),
        sg.Button('Merge', key='-MERGE-'),
        sg.Button('Insert', key='-INSERT-'),
        sg.Button('Replace', key='-REPLACE-'),
        sg.Button('Explode', key='-EXPLODE-')]])]
    return main_frame


def make_extract_layout():
    #  note tab group is expecting list[list] so this is double listed
    extract_layout = [
        [sg.Button('Open File', key='-EXTRACT-')],
        [sg.InputText(key='Extract Save As', do_not_clear=False, enable_events=True, visible=False),
         sg.FileSaveAs(key='-EXTRACTCURRENTPAGE-', file_types=(('PDF', '*.pdf'),))]]
    return extract_layout


def make_explode_layout():
    explode_layout = [
        [sg.Button('Open File', key='-EXPLODE-')],
        [sg.Button('Explode current document', key='-EXPLODEDOCUMENT-')]]
    return explode_layout


def make_delete_layout():
    delete_layout = [
        [sg.Button('Open File', key='-DELETE-')],
        [sg.Text('Delete from'), sg.InputText(key='-DELETEFROM-', size=6)
            , sg.Text('Delete to'), sg.InputText(key='-DELETETO-', size=6),
         sg.InputText(key='Delete Save As', do_not_clear=False, enable_events=True, visible=False),
         sg.FileSaveAs(key='-DELETEPAGES-', file_types=(('PDF', '*.pdf'),)), ]]
    return delete_layout


def make_merge_layout():
    merge_layout = [
        [sg.Input(key='-MERGE-'), sg.FilesBrowse(file_types=(('PDF', '*.pdf'),))],
        [sg.Button('Add')],
        [sg.Listbox(values=[''], key='-LISTBOXLIST-', size=(30, 6), expand_y=True, expand_x=True),
         sg.InputText(key='Merge Save As', do_not_clear=False, enable_events=True, visible=False),
         sg.FileSaveAs(key='-MERGEFILES-', file_types=(('PDF', '*.pdf'),)), ]]
    return merge_layout


def make_tab_group():
    tab_group_layout = [[sg.Tab('Extract', make_extract_layout(), key='-EXTRACTTAB-'),
                         sg.Tab('Explode', make_explode_layout(), key='-EXPLODETAB-'),
                         sg.Tab('Delete', make_delete_layout(), key='-DELETETAB-'),
                         sg.Tab('Merge', make_merge_layout(), key='-MERGETAB-'),
                         ]]
    return tab_group_layout


def make_window(theme):
    sg.theme(theme)
    menu_def = [['&Application', ['&Exit']],
                ['&Help', ['&About']]]
    right_click_menu_def = [[], ['Edit Me', 'Versions', 'Nothing', 'More Nothing', 'Exit']]

    # Table Data
    data = [["John", 10], ["Jen", 5]]
    headings = ["Name", "Score"]

    layout = [[sg.MenubarCustom(menu_def, background_color='white', text_color='black', key='-MENU-',
                                font=('Calibri', 10), tearoff=False)]]

    layout += [[[sg.TabGroup(make_tab_group(), enable_events=True, key='-TABGROUP-')], make_preview()]]

    layout[-1].append(sg.Sizegrip())
    window = sg.Window('PDF Editor', layout, right_click_menu=right_click_menu_def,
                       right_click_menu_tearoff=True, grab_anywhere=True, resizable=True, margins=(0, 0),
                       use_custom_titlebar=True, finalize=True, keep_on_top=True, location=(50, 50),
                       font=('Calibri', 10), titlebar_font=('Calibri', 11))
    window.set_min_size(size=(0, 0))
    return window


def setup_preview_window(mywindow):
    fname = sg.popup_get_file("Select file and filetype to open:", title="PyMuPDF Document Browser",
                              file_types=(("PDF Files", "*.pdf"),), keep_on_top=True, location=(200, 200),
                              font=('Calibri', 10))
    if fname is None or fname == '':  # user pressed X or Cancel
        return None, None, None, None, None
    else:
        pdfdocument, current_image, page_count, page_number = process_file(fname)
        mywindow.Element('-PAGENUMBER-').Update('1 of')
        mywindow.Element('-PREVIEW-').Update(data=current_image)
        mywindow.Element('-TotalPages-').Update(str(page_count) + ' Pages')
        return fname, pdfdocument, current_image, page_count, page_number


def do_extraction(fname, page_number, save_filename):
    if save_filename:
        print(save_filename)
        # p = pathlib.Path(fname)
        # stem = p.stem
        # directory = p.parent
        src_pdf = fitz.open(fname)
        out_pdf = fitz.open()
        out_pdf.insert_pdf(src_pdf, from_page=page_number, to_page=page_number)
        out_pdf.save(save_filename)
        sg.popup('File saved', keep_on_top=True)


def do_explosion(fname):
    basename = os.path.splitext(os.path.basename(fname))[0]
    directory = os.path.dirname(fname) + '/'

    src_pdf = fitz.open(fname)
    for page in range(len(src_pdf)):
        out_pdf = fitz.open()
        out_pdf.insert_pdf(src_pdf, from_page=page, to_page=page)
        out_pdf.save(directory + '{}_page_{}.pdf'.format(basename, page + 1))


def do_deletion(fname, page_from, page_to, save_filename):
    if not valid_page_range(fname, page_from, page_to):
        sg.popup("There is something wrong with the specified range", keep_on_top=True)
        print('invalid')
    else:
        page_from = int(page_from) - 1
        page_to = int(page_to) - 1
        src_pdf = fitz.open(fname)
        src_pdf.delete_pages(page_from, page_to)
        src_pdf.save(save_filename, garbage=3, clean=True, deflate=True)
        sg.popup('Edited file saved', keep_on_top=True)


def digits_only(string_to_test):
    '''
    :param string_to_test: value from window is a string, not an integer. Need to test if it is an integer.
    :return: True if the string only contains digits, false otherwise
    '''
    digits = '0123456789'
    for entry in string_to_test:
        if entry not in digits:
            return False
    return True


def valid_page_range(fname, page_from, page_to):
    src_pdf = fitz.open(fname)
    print(len(src_pdf))
    if not digits_only(page_from):
        return False  # not an integer
    elif not digits_only(page_to):
        return False  # not an integer
    elif page_to < page_from:
        return False  # delete_to must be >= delete_from
    elif int(page_from) <= 0 or int(page_to) <= 0:
        return False  # won't accept negative numbers and zero indexing
    elif int(page_to) > len(src_pdf):
        return False  # can't delete page number greater than document length
    else:
        return True


def add_files_to_merge_list(window, filelist):
    global merge_file_list
    for entry in filelist.split(';'):
        if entry not in merge_file_list:
            merge_file_list.append(entry)
    window.Element('-LISTBOXLIST-').Update(merge_file_list)


def main():
    window = make_window(sg.theme())

    # This is an Event Loop
    while True:
        window, event, values = sg.read_all_windows(timeout=100)
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
        elif event == "-EXTRACT-":
            print("[LOG] Open file for extract")
            fname, pdfdocument, current_image, page_count, page_number = setup_preview_window(window)
            print('fname is ', fname)
            window.Element('-PREVIEWFRAME-').Update(visible=True)
            print("[LOG] User chose file: " + str(fname))
        elif event == "Extract Save As":
            print("[LOG] Extracting a page")
            if 'fname' not in locals():
                sg.popup("You haven't opened a file yet", keep_on_top=True, font=('Calibri', 10))
            else:
                do_extraction(fname, page_number, values['-EXTRACTCURRENTPAGE-'])  # value is filename
        elif event == "-EXPLODE-":
            print("[LOG] Open file for explode")
            fname, pdfdocument, current_image, page_count, page_number = setup_preview_window(window)
            print('fname is ', fname)
            window.Element('-PREVIEWFRAME-').Update(visible=True)
            print("[LOG] User chose file: " + str(fname))
        elif event == "-EXPLODEDOCUMENT-":
            if 'fname' not in locals():
                sg.popup("You haven't opened a file yet", keep_on_top=True, font=('Calibri', 10))
            else:
                do_explosion(fname)
        elif event == "-DELETE-":
            print("[LOG] Open file for delete")
            fname, pdfdocument, current_image, page_count, page_number = setup_preview_window(window)
            print('fname is ', fname)
            window.Element('-PREVIEWFRAME-').Update(visible=True)
        elif event == "Delete Save As":
            print("[LOG] Deleting a page")
            page_from = values['-DELETEFROM-']
            page_to = values['-DELETETO-']
            if 'fname' not in locals():
                sg.popup("You haven't opened a file yet", keep_on_top=True, font=('Calibri', 10))
            else:
                do_deletion(fname, page_from, page_to, values['-DELETEPAGES-'])
        elif event == "Add":
            add_files_to_merge_list(window, values['-MERGE-'])
        elif event == "Next":
            if 'pdfdocument' in locals():
                if page_number < page_count - 1:  # have to wrap around
                    next_page = page_number + 1
                else:
                    next_page = 0
                new_data, page_number = get_page(pdfdocument, next_page)
                # print(new_data)
                window.Element('-PREVIEW-').Update(data=new_data)
                window.Element('-PAGENUMBER-').Update(str(page_number + 1) + ' of')
        elif event == "Prev":
            if 'pdfdocument' in locals():
                if page_number > 0:  # have to wrap around
                    next_page = page_number - 1
                else:
                    next_page = page_count - 1
                new_data, page_number = get_page(pdfdocument, next_page)
                # print(new_data)
                window.Element('-PREVIEW-').Update(data=new_data)
                window.Element('-PAGENUMBER-').Update(str(page_number + 1) + ' of')
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
    sg.theme('SystemDefault1')  # sg.theme('DarkGrey7')
    main()
