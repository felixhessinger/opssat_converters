'''
Author: Felix Hessinger
Email: felix.hessinger@gmail.com
Phone: +49-1578-7373424
Date: 17/06/2019

This script generates se.xml files needed for MATIS. In MATIS each System Element needs a se.xml file if minimum one
procedure is included. The se.xml file contains information about the content of the System Element. These contents
include procedure names, their configuration and their input arguments with description and ID.
Input:
"Excel" folder located inside the current workspace. It iterates through all sub-folders, reads in each Excel procedures
and extracts the needed information.
Output:
A se.xml file containing the name, configuration, ID and input arguments of all procedures inside each folder containing
Excel procedures.
The file name is the current sub-folder name plus .se.xml .

DISCLAIMER: It is just a tool to make life easier. Please check each generated file for errors before implementing it.
            This code might contain overseen bugs.
'''
import openpyxl as op  # use version 2.5.3, newer versions might not work
import os

def get_last_folder_name(_files):
    try:
        _file_path = _files.rsplit('\\')[1]
    except:
        _file_path = None
    return _file_path


def escape_special_characters(_procedure_description, _argument_description):
    """
    Escapes special characters in procedure and argument description and converts them into xml format.
    """
    _converted_procedure_description = str(_procedure_description).replace('&', '&amp;').replace('\'', '&apos;').replace('<', '&lt;').replace('>', '&gt;').replace('\"', '&quot;')
    _converted_argument_description = str(_argument_description).replace('&', '&amp;').replace('\'', '&apos;').replace('<', '&lt;').replace('>', '&gt;').replace('\"', '&quot;')
    return _converted_procedure_description, _converted_argument_description


def get_procedure_name_and_description(_names_of_current_level, _folder_level):
    """
    Extracts the procedure name and its description from the file name and Returns these as string.
    """
    try:
        _procedure_name = files.rsplit('\\', 1)[0]
        file_path = _names_of_current_level.rsplit('\\')[1]
    except:
        print('', end='')
		
    a = _names_of_current_level.find("_")
    _procedure_name = _names_of_current_level[0:a]
    try:
        _procedure_description = _names_of_current_level[a::].rsplit('.xlsx', 1)[0].replace('_', ' ')
        if _procedure_description.startswith(' '):
            _procedure_description = _procedure_description[1::]
    except:
        print('', end='')
    return _procedure_name, _procedure_description
    

def get_flags(_current_last_folder_name, _past_last_folder_name, _list_of_excelsheet_paths, _counter_nth_file):
    if _past_last_folder_name != _current_last_folder_name:
        _flag_xml_new_file = 1
    else:
        _flag_xml_new_file = 0
    try:
        _future_last_folder_name = get_last_folder_name(_list_of_excelsheet_paths[_counter_nth_file + 1])
    except:
        _future_last_folder_name = 'None'

    if _future_last_folder_name != _current_last_folder_name or _future_last_folder_name == 'None':
        _flag_end_of_xml_file = 1
    else:
        _flag_end_of_xml_file = 0
    return _flag_xml_new_file, _flag_end_of_xml_file


def get_argument_name_and_description(_file_with_path):
    """
    Returns argument ID, Description and Type extracted from the corresponding Excel file.
    :param _file_with_path: complete file path
    :return: argument ID, Description and Type
    """
    print(_file_with_path)
    wb = op.load_workbook(_file_with_path, data_only=True)
    ws_procedure = wb['Procedure']
    arguments_ID, arguments_DESCRIPTION, arguments_TYPE = [], [], []
    row_number = 0
    for row in ws_procedure['{START_LETTER}1:{STOP_LETTER}20'.format(START_LETTER=_OPERATIONS_COLUMN, STOP_LETTER=_OPERATIONS_COLUMN)]:
        row_number += 1
        for cell in row:
            if str(cell.internal_value).startswith('Parameters:'):
                row_number_parameter_start = row_number
                break
        else:
            continue
        break
    row_number_current = row_number_parameter_start
    while ws_procedure['{ID_COLUMN}{ROW_NUMBER_CURRENT}'.format(ID_COLUMN=_ID_COLUMN, ROW_NUMBER_CURRENT=row_number_current)].internal_value != None:
        arguments_ID.append(str(ws_procedure['{ID_COLUMN}{ROW_NUMBER_CURRENT}'.format(ID_COLUMN=_ID_COLUMN, ROW_NUMBER_CURRENT=row_number_current)].internal_value).replace('$', ''))
        arguments_DESCRIPTION.append(ws_procedure['{DESCRIPTION_COLUMN}{ROW_NUMBER_CURRENT}'.format(DESCRIPTION_COLUMN=_DESCRIPTION_COLUMN, ROW_NUMBER_CURRENT=row_number_current)].internal_value)
        arguments_TYPE.append(ws_procedure['{TYPE_COLUMN}{ROW_NUMBER_CURRENT}'.format(TYPE_COLUMN=_TYPE_COLUMN, ROW_NUMBER_CURRENT=row_number_current)].internal_value)
        row_number_current += 1
    # print(arguments_ID, arguments_DESCRIPTION, arguments_TYPE)
    return arguments_ID, arguments_DESCRIPTION, arguments_TYPE


def write_seXml_files(_current_last_folder_name, _paths_without_file, _procedure_name, _procedure_description, _flag_new_file, _flag_last_procedure_in_folder, _arguments_name, _arguments_description, _arguments_type, _counter_nth_file):
    file_output_path = str(_paths_without_file[_counter_nth_file]).replace('\\Excel\\', '\\generated_MATIS_Files\\')
    fo = open(file_output_path + '\\' + _current_last_folder_name + '.se.xml', 'a+')
    if _flag_new_file:
        fo.close()
        try:
            os.remove(str(file_output_path) + '\\' + _current_last_folder_name + '.se.xml')
        except:
            print('', end='')
        fo = open(file_output_path + '\\' + _current_last_folder_name + '.se.xml', 'a+')
        fo.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<LocalSystemElement SchemaVersion="1.0">\n')
    fo.write('\t<SEObject Name="{PROCEDURE_NAME}">\n'.format(PROCEDURE_NAME=_procedure_name))
    _procedure_description, _ = escape_special_characters(_procedure_description, _arguments_description)
    fo.write(
        '\t\t<ActivityDefinition Description="{PROCEDURE_DESCRIPTION}" Constraints="" Objectives="" Preconditions="" Postconditions="" EstimatedDuration="000:00:05:00.000" ValidationState="draft">\n'.format(
            PROCEDURE_DESCRIPTION=_procedure_description))
    for_counter = 0
    for _argument_name,_argument_description,_argument_type in zip(_arguments_name,_arguments_description, _arguments_type):
        _, _argument_description = escape_special_characters(_procedure_description, _argument_description)
        _argument_type = convert_TYPE_from_SCOS_to_MATIS(_argument_type)
        fo.write(
            '\t\t\t<Argument Name="{ARGUMENT_NAME}" Description="{ARGUMENT_DESCRIPTION}">\n\t\t\t\t<Scalar Type="{ARGUMENT_TYPE}"/>\n\t\t\t</Argument>\n'.format(
                ARGUMENT_NAME=_argument_name, ARGUMENT_TYPE=_argument_type, ARGUMENT_DESCRIPTION=_argument_description))
        for_counter += 1
    fo.write('\t\t</ActivityDefinition>\n')
    fo.write('\t</SEObject>\n')
    if _flag_last_procedure_in_folder:
        fo.write('</LocalSystemElement>\n')
    fo.close()
    return


def convert_TYPE_from_SCOS_to_MATIS(current_TYPE_cell_value):
    if str(current_TYPE_cell_value).startswith('Enum') or str(current_TYPE_cell_value).startswith('U8') or str(current_TYPE_cell_value).startswith('U16') or str(current_TYPE_cell_value).startswith('U32') or str(current_TYPE_cell_value).startswith('U64'):
        return 'unsignedInteger'
    elif str(current_TYPE_cell_value).startswith('S8') or str(current_TYPE_cell_value).startswith('S16') or str(current_TYPE_cell_value).startswith('S32') or str(current_TYPE_cell_value).startswith('S64'):
        return 'signedInteger'
    elif str(current_TYPE_cell_value).startswith('Boolean'):
        return 'boolean'
    elif str(current_TYPE_cell_value).startswith('Float'):
        return 'real'
    elif str(current_TYPE_cell_value).startswith('Octet Str') or str(current_TYPE_cell_value).startswith('Char Str'):
        return 'string'
    elif str(current_TYPE_cell_value).startswith('Abs Time') or str(current_TYPE_cell_value).startswith('Abs time'):
        return 'absoluteTime'
    elif str(current_TYPE_cell_value).startswith('Del Time') or str(current_TYPE_cell_value).startswith('Del time'):
        return 'relativeTime'
#TODO: change this to something that is working but unique, so that MATIS does not complain but user can identify "wrong" value type
    return 'string'


##############################
# Definitions of global variables
_STEP_COLUMN = 'A'
_OPERATIONS_COLUMN = 'B'
_ID_COLUMN = 'C'
_DESCRIPTION_COLUMN = "D"
_TYPE_COLUMN = "E"
##############################
print('Converter started')
rootdir = os.getcwd()
list_of_excelsheet_paths =[]
files_with_path = []
paths_without_file = []
for subdir, dirs, files in os.walk(rootdir + '\\Excel\\'):
    for file in files:
        #print os.path.join(subdir, file)
        filepath = subdir + os.sep + file
        if filepath.endswith(".xlsx") and '\\old\\' not in filepath and '~' not in filepath:
            files_with_path.append(str(filepath))
            try:
                paths_without_file.append(str(filepath).rsplit('\\', 1)[0])
            except:
                print('', end='')
            list_of_excelsheet_paths.append(str(filepath).split('\\Excel\\', 1)[1])
# print(paths_without_file)
# print(files_with_path)
# print(list_of_excelsheet_paths)

names_of_current_level = []
the_end = ['not empty', 'even less empty']
#list_of_excelsheet_paths.append('asdf.xlsx')
folder_level = 0
counter_nth_file = 0
flag_xml_new_file = 1
past_last_folder_name = ''
# for file_with_path in files_with_path:
#     get_argument_name_and_description(file_with_path)
print('Creating files...\n(This might take a minute)')
while len(the_end) > 0:
    the_end = []
    procedure_name, procedure_description = 'None', 'None'
    for file_with_half_path, file_with_complete_path in zip(list_of_excelsheet_paths, files_with_path):
        #print(file_with_half_path)
        #print(file_with_complete_path)
        try:
            names_of_current_level = file_with_half_path.split('\\')[folder_level]
            the_end.append(names_of_current_level)
        except:
            continue
        if ('.xlsx' in names_of_current_level):
            #print(names_of_current_level)
            #print('asfd' + str(folder_level))
            current_last_folder_name = get_last_folder_name(file_with_half_path)
            flag_xml_new_file, flag_end_of_xml_file = get_flags(current_last_folder_name, past_last_folder_name, list_of_excelsheet_paths, counter_nth_file)
            procedure_name, procedure_description = get_procedure_name_and_description(names_of_current_level, folder_level)
            procedure_name = str(procedure_name).replace('-', '_')
            arguments_ID, arguments_DESCRIPTION, arguments_TYPE = get_argument_name_and_description(file_with_complete_path)
            write_seXml_files(current_last_folder_name, paths_without_file, procedure_name, procedure_description, flag_xml_new_file, flag_end_of_xml_file, arguments_ID, arguments_DESCRIPTION, arguments_TYPE, counter_nth_file)
            past_last_folder_name = current_last_folder_name
            counter_nth_file += 1
    folder_level += 1
print('Files created. Have fun! :)')
