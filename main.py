import PySimpleGUI as sg
import pandas as pd
import numpy as np
import sys
import xlwings as xw


def _get_excel_heading_list(excel_file_path: str) -> pd.DataFrame.columns:
    headings_list = pd.read_excel(io=f'{excel_file_path}', nrows=0)
    return headings_list.columns

def _get_matching_elements(array1: np.array, array2: np.array) -> list[tuple[int, int]]:
    matching_list = []
    for i in range(len(array1)):
        for j in range(len(array2)):
            if str(array1[i]) == str(array2[j]):
                matching_list.append((i, j))
    return matching_list


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
            headings = _get_excel_heading_list(sys_msng_path)           
            sys_msng_data = pd.read_excel(io=f'{sys_msng_path}', usecols=[0,1,2,3], dtype={f'{headings[1]}': float})
            sys_msng_part_numbers = sys_msng_data[f'{headings[0]}'].values
            sys_msng_qty_missing = sys_msng_data[f'{headings[1]}'].values
            sys_msng_location = sys_msng_data[f'{headings[2]}'].values
            sys_msng_date_updated = sys_msng_data[f'{headings[3]}'].values

            phys_msng_path = values[1]
            headings = _get_excel_heading_list(phys_msng_path)           
            phys_msng_data = pd.read_excel(io=f'{phys_msng_path}', usecols=[0,1,2,3], dtype={f'{headings[1]}': float})
            phys_msng_part_numbers = phys_msng_data[f'{headings[0]}'].values
            phys_msng_qty_missing = phys_msng_data[f'{headings[1]}'].values
            phys_msng_location = phys_msng_data[f'{headings[2]}'].values
            phys_msng_date_updated = phys_msng_data[f'{headings[3]}'].values

            matching_part_numbers_index = _get_matching_elements(sys_msng_part_numbers, phys_msng_part_numbers)
            matching_qty_missing_index = _get_matching_elements(sys_msng_qty_missing, phys_msng_qty_missing)
            matching_parts_index = _get_matching_elements(matching_part_numbers_index, matching_qty_missing_index)

            
            matching_parts = {}
            for (i, j) in matching_parts_index:
                matching_parts[phys_msng_part_numbers[matching_part_numbers_index[i][1]]] = [phys_msng_qty_missing[matching_qty_missing_index[j][1]], 
                                                                                             phys_msng_location[matching_part_numbers_index[i][1]], 
                                                                                             sys_msng_location[matching_part_numbers_index[i][0]]]
            window.close()
            break

        if event == 'Cancel' or event == sg.WIN_CLOSED:
            window.close()
            sys.exit()
    

    window_layout = [[sg.Text('List of fully matching inventory errors:')]]
    for element in matching_parts:
        window_layout.append([sg.Text(f'{element}: {matching_parts[element][0]:g}'), 
                              sg.Push(),
                              sg.Text(f'{matching_parts[element][1]} --> {matching_parts[element][2]}'), 
                              sg.Button('Commit Change')])
        
    window_layout.append([sg.Button('Commit All'), sg.Button('Exit')])

    window = sg.Window('Inventory Management System', window_layout)

    while True:
        event, values = window.read()
        if event == 'Exit' or event == sg.WIN_CLOSED:
            window.close()
            sys.exit()
    
