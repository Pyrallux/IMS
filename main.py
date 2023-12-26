import PySimpleGUI as sg
import pandas as pd
import numpy as np
import sys
import xlwings as xw

def _get_excel_heading_list(excel_file_path: str) -> pd.DataFrame.columns:
    headings_list = pd.read_excel(io=f'{excel_file_path}', nrows=0)
    return headings_list.columns

def _get_matching_elements(array1: pd.array, array2: pd.array) -> list[tuple[int, int]]:
    matching_list = []
    for i in range(len(array1)):
        for j in range(len(array2)):
            if str(array1[i]) == str(array2[j]):
                matching_list.append((i, j))
    return matching_list

def _index_data(excel_data_path):
    app = xw.App(visible=False)
    wb = xw.Book(f'{excel_data_path}')
    sht = wb.sheets['Sheet1']
    data_range = sht.tables[0].data_body_range
    data = sht.range(data_range).value
    part_numbers = [data[i][0] for i in range(len(data))]
    for j in range(len(part_numbers)):
        if type(part_numbers[j]) != str:
            part_numbers[j] = f'{int(part_numbers[j])}'
    qty_missing = [float(data[i][1]) for i in range(len(data))]
    location = [str(data[i][2]) for i in range(len(data))]
    wb.close()
    return data, part_numbers, qty_missing, location


if __name__ == '__main__':

    # Setup initial window with file browsing for two part lists
    sg.theme('Dark2')
    window_layout = [[sg.Text('Welcome to the Inventory Management System')],
            [sg.Text('Filepath to Systematically Missing Parts')],
            [sg.Input(), sg.FileBrowse()],
            [sg.Text('Filepath to Physically Missing Parts')],
            [sg.Input(), sg.FileBrowse()],
            [sg.Button('Next'), sg.Button('Cancel')]]
    window = sg.Window('Inventory Management System', window_layout)

    # Event loop waiting for button activity
    while True:
        event, values = window.read()
        if event == 'Next':
            sys_msng_path = values[0]
            phys_msng_path = values[1]
            sys_msng_data, sys_msng_part_numbers, sys_msng_qty_missing, sys_msng_location = _index_data(sys_msng_path)
            phys_msng_data, phys_msng_part_numbers, phys_msng_qty_missing, phys_msng_location = _index_data(phys_msng_path)

            matching_part_numbers_index = _get_matching_elements(sys_msng_part_numbers, phys_msng_part_numbers)
            matching_qty_missing_index = _get_matching_elements(sys_msng_qty_missing, phys_msng_qty_missing)
            matching_parts_index = _get_matching_elements(matching_part_numbers_index, matching_qty_missing_index)

            matching_parts = {}
            n = 0
            for (i, j) in matching_parts_index:
                matching_parts[n] = [phys_msng_part_numbers[matching_part_numbers_index[i][1]],
                                     phys_msng_qty_missing[matching_qty_missing_index[j][1]], 
                                     phys_msng_location[matching_part_numbers_index[i][1]], 
                                     sys_msng_location[matching_part_numbers_index[i][0]]]
                n += 1
            window.close()
            break

        if event == 'Cancel' or event == sg.WIN_CLOSED:
            window.close()
            sys.exit()
    

    window_layout = [[sg.Text('List of fully matching inventory errors:')]]
    for element in matching_parts:
        window_layout.append([sg.Text(f'{matching_parts[element][0]}: {matching_parts[element][1]:g}'), 
                              sg.Push(),
                              sg.Text(f'{matching_parts[element][2]} --> {matching_parts[element][3]}'), 
                              sg.Button('Commit Change')])
    window_layout.append([sg.Button('Commit All'), sg.Button('Exit')])
    window = sg.Window('Inventory Management System', window_layout)

    while True:
        sys_msng_removed_parts_list = []
        phys_msng_removed_parts_list = []
        event, values = window.read()
        if 'Commit Change' in event:
            if event == 'Commit Change':
                element_commited_index = 0
            else:
                element_commited_index = int(event[13:]) + 1
            element_committed = matching_part_numbers_index[matching_parts_index[element_commited_index][0]]

            app = xw.App(visible=False)

            sys_msng_wb = xw.Book(f'{sys_msng_path}')
            sys_msng_sht = sys_msng_wb.sheets['Sheet1']

            phys_msng_wb = xw.Book(f'{phys_msng_path}')
            phys_msng_sht = phys_msng_wb.sheets['Sheet1']
            
            del sys_msng_data[element_committed[0]]

            print(sys_msng_sht.range(f'A2:D{len(sys_msng_data) + 1}').value)
            sys_msng_sht.tables[0].data_body_range.clear()
            sys_msng_sht.range(f'A2:D{len(sys_msng_data) + 1}').value = sys_msng_data
            print(sys_msng_sht.range(f'A2:D{len(sys_msng_data) + 1}').value)

            sys_msng_wb.save(sys_msng_path)
            sys_msng_wb.close()
            phys_msng_wb.close()
            app.quit()

        if event == 'Exit' or event == sg.WIN_CLOSED:
            window.close()
            sys.exit()

    
    
