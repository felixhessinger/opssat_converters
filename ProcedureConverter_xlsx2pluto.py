'''
Author: Felix Hessinger
Email: felix.hessinger@gmail.com
Phone: +49-1578-7373424
Date: 03/07/2019

INFO: For questions, comments or improvements regarding the converter feel free to contact me.

DESCRIPTION:
This script converts a Excel procedure with a Standard defined by OPS-SAT into the PLUTO language for easier integration
to MATIS.
Input:
"Excel" folder located inside the current workspace. It iterates through all sub-folders, reads-in each Excel procedures
and extracts the needed information.
Output:
A representative PLUTO procedure file to the Excel equivalent in the correct folder structure.
The file name is the Excel file name without its description.

DISCLAIMER: It is just a tool to make life easier. Please check each generated file for errors before implementing it.
            This code might contain overseen bugs.
'''

import openpyxl as op  # use version 2.5.3, newer versions might not work
import fileinput
import re
import os
import shutil
import time
import datetime


# FUNCTION DEFINITIONS
def write_DATE_of_autogeneration_and_initials():
    '''
    Writes information about generation time, information about converter,
    contact information and disclaimer into the output file.

    :return:
    '''
    f = open(_FILE_NAME, 'w+')
    f.write('////////////////////////////////////////////////////////////////////////////////////////\n')
    f.write('// Date for Base Code auto-generation: ' + str(datetime.datetime.now()) + '\n')
    f.write('// Converter designed by: Felix Tim Hessinger\n')
    f.write('//\n')
    f.write('// Last manually edited at: None\n')
    f.write('// Last manually edited by: None\n')
    f.write('//\n')
    f.write('// Last tested and validated at: None\n')
    f.write('// Last tested and validated by: None\n')
    f.write('//\n')
    f.write('// INFO: For questions, comments or improvements regarding the converter feel free to contact me.\n')
    f.write('// Email: felix.hessinger@gmail.com             Phone: (+49)1578-7373424\n')
    f.write('// DISCLAIMER: The Converter used is just a tool to make conversion easier, it might contain bugs or wrong syntax\n')
    f.write('//             for certain cases, so make sure to double-check the generated PLUTO code with procedure itself.\n')
    f.write('////////////////////////////////////////////////////////////////////////////////////////\n')
    f.write('////////////////////////////////////////////////////////////////////////////////////////\n')
    f.close()
    return


def write_front_page_documentation_as_comment_into_f(ws_front_page):
    '''
    Writes the front page of the current Excel procedure into
    the output file as a comment.

    :param ws_front_page:
    :return:
    '''
    # open dump-f for generated Code
    f = open(_FILE_NAME, 'a+')
    end_of_excel_sheet = False
    old_cell_value = None
    counter_row = 1
    counter = 0
    while not end_of_excel_sheet:
        current_cell_value = ws_front_page[_INFORMATION_COLUMN_FRONT_PAGE + '{row}'.format(row=counter_row)].value
        #print(current_cell_value)
        if current_cell_value == old_cell_value:
            counter += 1
        else:
            counter = 0
        if counter >= 5:
            end_of_excel_sheet = True
            counter_row -= 4
        counter_row += 1
        old_cell_value = current_cell_value
    for cell_number in range(1, counter_row):
        f.write('//')
        for cell_ID in [_START_COLUMN_FRONT_PAGE, _FILLER_COLUMN1_FRONT_PAGE, _FILLER_COLUMN2_FRONT_PAGE, _INFORMATION_COLUMN_FRONT_PAGE, _END_COLUMN_FRONT_PAGE]:  # loop for column A to E (start to end)
            if ws_front_page[cell_ID + '{row}'.format(row=cell_number)].value != None:
                f.write(ws_front_page[cell_ID + '{row}'.format(row=cell_number)].value)
            else:
                f.write('\t\t\t\t\t')
        f.write('\n')
    f.close()
    check_front_page_for_errors()
    f = open(_FILE_NAME, 'a+')
    f.write('////////////////////////////////////////////////////////////////////////////////////////\n\n')
    f.write('////////////////////////////////////////////////////////////////////////////////////////\n')
    f.write('// START OF PROCEDURE CODE\n')
    f.write('////////////////////////////////////////////////////////////////////////////////////////\n')
    f.close()
    return


def check_front_page_for_errors():
    '''
    Checks for uncommented lines in input files.
    Usually used after writing the front page into the file and
    checking for errors due to returns in an input Excel cell.

    :return:
    '''
    for line in fileinput.FileInput(_FILE_NAME, inplace=1):
        if not line.startswith('//'):
            line = line.replace(line, '//\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t'+line)
            print(line, end='')
        else:
            print(line, end='')
    return


def check_file_for_forbidden_characters(_FILE_NAME):
    '''
    Checks file for non-ASCII characters and deletes them.

    :param _FILE_NAME:
    :return:
    '''
    for line in fileinput.FileInput(_FILE_NAME, inplace=1):
        for char in line:
            if not char in 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ !\"#$%&\'()*+,-~./0123456789:;<=>?@[\\]^_\t\n`':
                char = char.replace(char, '')
                print(char, end='')
            else:
                print(char, end='')
    return


def check_file_for_empty_steps_and_delete(_FILE_NAME):
    '''
    Checks file for empty steps and deletes them.

    :param _FILE_NAME:
    :return:
    '''
    old_old_line, old_line = '', ''
    for current_line in fileinput.FileInput(_FILE_NAME, inplace=1):
        current_line_ = current_line.replace(' ', '')
        if not (str(old_line).replace(' ', '').replace('\t', '').startswith('initiateandconfirmstep') and str(current_line).replace(' ', '').replace('\t', '').startswith('endstep')) and not (str(old_old_line).replace(' ', '').replace('\t', '').startswith('initiateandconfirmstep') and str(old_line).replace(' ', '').replace('\t', '').startswith('endstep')):
            print(old_line, end='')
        old_old_line = old_line
        old_line = current_line
    f = open(_FILE_NAME, 'a')
    f.write(current_line)
    f.close()
    return


def get_operations_captions_row_number(ws_procedure):
    '''
    Gets operations caption row number.

    :param ws_procedure:
    :return newOperations:
    '''
    for i in [_STEP_COLUMN]:  # a sequence of columns could be initialised with: ['A','B','C','D','E','F','G']
        newOperations = []
#TODO: change 10001 to end of excel file (low priority)
        for j in range(1, 10001):
            currentCell = '{col}{row}'.format(col=i, row=j)
            cellColor = ws_procedure[currentCell].fill.start_color.index
            if (cellColor == _COLOR_DIVIDING_OPERATION_STEPS):       #blue-ish
                newOperations.append(j)
    return newOperations


def iterating_over_operation_topic(ws_procedure, new_operation_row_numbers, iterationNumber, identifier_matrix):
    '''
    iterates over the operation topics and decides if it is a FOLLOW_ID_FIELD, NEW_ID_FIELD, NEW_OPERATION_FIELD or FOLLOW_OPERATION_FIELD.

    :param ws_procedure, new_operation_row_numbers, iterationNumber, identifier_matrix:
    :return identifier_matrix:
    '''
    flag_previous_ID_empty = None
    old_OPERATIONS_cell_value = None
    for rows_between_two_operations_topics in range((new_operation_row_numbers[iterationNumber]-1)+2, new_operation_row_numbers[iterationNumber+1]):
#TODO: change _ID_COLUMN to _OPERATIONS_COLUMN to be able to get all the commands that have to be written into MATIS code. E.g. between PREPARATION and INITIAL_VERIFICATIONS
        current_ID_cell_value = ws_procedure[_ID_COLUMN + '{rows_procedure}'.format(rows_procedure=rows_between_two_operations_topics)].value
        current_OPERATIONS_cell_value = ws_procedure[_OPERATIONS_COLUMN + '{rows_procedure}'.format(rows_procedure=rows_between_two_operations_topics)].value
        next_OPERATIONS_cell_value = ws_procedure[_OPERATIONS_COLUMN + '{rows_procedure}'.format(rows_procedure=rows_between_two_operations_topics)].value
        if current_ID_cell_value == None:
            flag_current_ID_empty = 1
        else:
            flag_current_ID_empty = 0

        if current_OPERATIONS_cell_value == None:
            flag_current_OPERATIONS_empty = 1
        else:
            flag_current_OPERATIONS_empty = 0
        # check if current cell is not empty
        if flag_current_ID_empty == 0:
            # check if previous cell was empty
            if flag_previous_ID_empty:
                # mark with "a" that this is the beginning of a clustered parameter block
                identifier_matrix.append([rows_between_two_operations_topics, 'NEW_ID_FIELD', str(current_OPERATIONS_cell_value)])
                old_OPERATIONS_cell_value = current_OPERATIONS_cell_value
            #check if previous cell was not empty
            elif flag_previous_ID_empty == 0:
                # mark with "b" that this is a following parameter
                identifier_matrix.append([rows_between_two_operations_topics, 'FOLLOW_ID_FIELD', str(old_OPERATIONS_cell_value) + '_'])
        elif flag_current_OPERATIONS_empty == 0:
            if str(current_OPERATIONS_cell_value).startswith(tuple(_KNOWN_OPERATIONS_PARAMETER)) or ws_procedure:
                identifier_matrix.append([rows_between_two_operations_topics, 'NEW_OPERATION_FIELD', str(current_OPERATIONS_cell_value)])
            else:
                identifier_matrix.append([rows_between_two_operations_topics, 'FOLLOW_OPERATION_FIELD', str(current_OPERATIONS_cell_value)])
        # update the flag for the next iteration
        flag_previous_ID_empty = flag_current_ID_empty
    return identifier_matrix


def create_parameter_dictionary():
    '''
    Creates parameter dictionary as a library in tree structure
    :param :
    :return tree_structure_params_repository:
    '''
    MIB_TCs = {'AggregationDefinitions': ['MG_'],
                'CCSDSengineHWTC': ['CCSDSTC1'],
                'FBO': ['F1E1'],
                'GroundSegment': ['G2A0101s'],
                'Misc': ['CFDPTC0'],
                'NanomindTCs_critical': ['M4A0B01b', 'M4A0B04b', 'M4A0E01b', 'M4A1301b', 'M4A1303b',
                                         'M4A1304b', 'M4B1601b', 'M4B1602b', 'M4B1604b', 'M4B1704i',
                                         'M4B1705i', 'M4B1708b'],
                'NanomindTCs': ['M040', 'M4A0', 'M4B1', 'MAx1']
                }

    MIB_TMs = {'ADC': ['ADCS', 'BUSS1137', 'CADC', 'GNC_', 'MAGN'],
                'ADF': ['IAC', 'BUSS1154'],
                'CAM': ['CAM'],
                'CCS': ['CCS', 'MF0000'],
                'DeviceCheckFlags': ['c'],
                'DHS': ['TM_', 'POCK', 'PACK', 'PP00'],
                'EPS': ['EPS', 'SOL'],
                'EXP': ['EXPE'],
                'GPS': ['GPS'],
                'OBC': ['INIT', 'NAN', 'OBSW', 'RTC', 'SW_M', 'SWMS', 'TIME'],
                'ORX': ['BUSS1140', 'ORX', 'BUSS1155'],
                'SDR': ['BUSS1139', 'SDR', 'BUSS1156'],
                'SEP': ['BUSS1138', 'SEP', 'BUSS1157'],
                'SXV': ['SBD'],
                'UHF': ['COM', 'UHF'],
                'XTX': ['XBD']
                }
    # creating tree structure of repository
    telecommands = {'MIB_TCs': MIB_TCs}
    telemetry = {'MIB_TMs': MIB_TMs}
    SSM = {'Telecommands': telecommands, 'Telemetry': telemetry}
    tree_structure_params_repository = {'SSM': SSM}

# VISUALISATION: check if all the MIB Parameter are in the correct folder in the Repository in MATIS (for drag and drop)
#     # print all Telecommands starts with their corresponding category
#     print('\n\nMIB_TCs')
#     for params in MIB_TCs:
#         print('\t' + params)
#         for pr in MIB_TCs['{par}'.format(par=params)]:
#             print('\t\t' + pr)
#         print()
#     # print all Telemetry starts with their corresponding category
#     print('\n\nMIB_TMs')
#     for params in MIB_TMs:
#         print('\t' + params)
#         for pr in MIB_TMs['{par}'.format(par=params)]:
#             print('\t\t' + pr)
#         print()
    return tree_structure_params_repository


def create_PROCEDURE_dictionary():
    '''
    Creates procedure dictionary as a library in tree structure
    :param :
    :return tree_structure_PROCEDURE_repository:
    '''
    decommissioning_nominal = {'SYS': ['DEC_SYS_N100']}
    LEOP_contingency = {'EPS': ['LEOP_EPS_C210'],
                'UHF': ['LEOP_UHF_C210'],
                }
    LEOP_nominal = {'ADC': ['LEOP_ADC_N100','LEOP_ADC_N200','LEOP_ADC_N250','LEOP_ADC_N300','LEOP_ADC_N400'],
                'ADF': ['LEOP_ADF_N100'],
                'CAM': ['LEOP_CAM_N100'],
                'CCS': ['LEOP_CCS_N100'],
                'DHS': ['LEOP_DHS_N200'],
                'EPS': ['LEOP_EPS_N100', 'LEOP_EPS_N200', 'LEOP_EPS_N300'],
                'EXP': ['LEOP_EXP_N100'],
                'FIL': ['LEOP_FIL_N100'],
                'GPS': ['LEOP_GPS_N100'],
                'OBC': ['LEOP_OBC_N100', 'LEOP_OBC_N200'],
                'ORX': ['LEOP_ORX_N100'],
                'SDR': ['LEOP_SDR_N100'],
                'SEP': ['LEOP_SEP_N100'],
                'SXV': ['LEOP_SXV_N100', 'LEOP_SXV_N200'],
                'SYS': ['LEOP_SYS_N200', 'LEOP_SYS_N300', 'LEOP_SYS_N350'],
                'UHF': ['LEOP_UHF_N100', 'LEOP_UHF_N200'],
                'XTX': ['LEOP_XTX_N100']
               }

    routine_contingency = {'ADC': ['R_ADC_C120', 'R_ADC_C310'],
                'CCS': ['R_CCS_C130'],
                'DHS': ['R_DHS_C100', 'R_DHS_C110', 'R_DHS_C410'],
                'EPS': ['R_EPS_C110', 'R_EPS_C120', 'R_EPS_C130', 'R_EPS_C150', 'R_EPS_C160', 'R_EPS_C170', 'R_EPS_C210', 'R_EPS_C220', 'R_EPS_C230', 'R_EPS_C310', 'R_EPS_C320', 'R_EPS_C410', 'R_EPS_C420', 'R_EPS_C430', ],
                'FIL': ['R_FIL_C200'],
                'GPS': ['R_GPS_C120', 'R_GPS_C210', 'R_GPS_C220', 'R_GPS_C310'],
                'OBC': ['R_OBC_C100', 'R_OBC_C150', 'R_OBC_C220', 'R_OBC_C250'],
                'SEP': ['R_SEP_C220', 'R_SEP_C230'],
                'SYS': ['R_SYS_C110', 'R_SYS_C120', 'R_SYS_C210', 'R_SYS_C220', 'R_SYS_C240', 'R_SYS_C310', 'R_SYS_C320', 'R_SYS_C410'],
                'UHF': ['R_UHF_C110', 'R_UHF_C120', 'R_UHF_C210', 'R_UHF_C220', 'R_UHF_C410', 'R_UHF_C510'],
                }
#TODO: add parameters to parameter_dictionary_TM
    routine_nominal = {'ADC': ['R_ADC_N110', 'R_ADC_N120', 'R_ADC_N210', 'R_ADC_N220', 'R_ADC_N230', 'R_ADC_N240', 'R_ADC_N310', 'R_ADC_N320', 'R_ADC_N510', 'R_ADC_N520'],
                'ADF': ['R_ADF_N110', 'R_ADF_N180', 'R_ADF_N210', 'R_ADF_N220', 'R_ADF_N230', 'R_ADF_N235', 'R_ADF_N240', 'R_ADF_N250', 'R_ADF_N260', 'R_ADF_N270', 'R_ADF_N280', 'R_ADF_N285', 'R_ADF_N350', 'R_ADF_N610'],
                'CAM': ['R_CAM_N110', 'R_CAM_N180','R_CAM_N210', 'R_CAM_N220', 'R_CAM_N350', 'R_CAM_N610'],
                'CCS': ['R_CCS_N110', 'R_CCS_N180', 'R_CCS_N210', 'R_CCS_N220', 'R_CCS_N225', 'R_CCS_N230', 'R_CCS_N240', 'R_CCS_N350', 'R_CCS_N410', 'R_CCS_N420', 'R_CCS_N425', 'R_CCS_N430', 'R_CCS_N450', 'R_CCS_N460', 'R_CCS_N610'],
                'DHS': ['R_DHS_N110', 'R_DHS_N210', 'R_DHS_N215', 'R_DHS_N410', 'R_DHS_N420', 'R_DHS_N425', 'R_DHS_N450', 'R_DHS_N510', 'R_DHS_N520'],
                'EPS': ['R_EPS_N110', 'R_EPS_N112', 'R_EPS_N120', 'R_EPS_N122', 'R_EPS_N125', 'R_EPS_N127', 'R_EPS_N130', 'R_EPS_N132', 'R_EPS_N135', 'R_EPS_N137', 'R_EPS_N140', 'R_EPS_N150', 'R_EPS_N160', 'R_EPS_N180', 'R_EPS_N350', 'R_EPS_N352', 'R_EPS_N510'],
                'EXP': ['R_EXP_N110', 'R_EXP_N120', 'R_EXP_N130', 'R_EXP_N210', 'R_EXP_N220', 'R_EXP_N230', 'R_EXP_N310', 'R_EXP_N320'],
                'FIL': ['R_FIL_N110', 'R_FIL_N150', 'R_FIL_N210', 'R_FIL_N220', 'R_FIL_N310', 'R_FIL_N320', 'R_FIL_N330', 'R_FIL_N340'],
                'GPS': ['R_GPS_N110', 'R_GPS_N120', 'R_GPS_N130', 'R_GPS_N210', 'R_GPS_N220', 'R_GPS_N510', 'R_GPS_N610'],
                'OBC': ['R_OBC_N110', 'R_OBC_N150', 'R_OBC_N180', 'R_OBC_N210', 'R_OBC_N350', 'R_OBC_N420', 'R_OBC_N425', 'R_OBC_N510', 'R_OBC_N522', 'R_OBC_N525', 'R_OBC_N527', 'R_OBC_N550', 'R_OBC_N555', 'R_OBC_N558', 'R_OBC_N560', 'R_OBC_N570'],
                'ORX': ['R_ORX_N110', 'R_ORX_N180', 'R_ORX_N210', 'R_ORX_N220', 'R_ORX_N350', 'R_ORX_N610'],
                'SDR': ['R_SDR_N110', 'R_SDR_N180', 'R_SDR_N210', 'R_SDR_N220', 'R_SDR_N350', 'R_SDR_N610'],
                'SEP': ['R_SEP_N110', 'R_SEP_N150', 'R_SEP_N180', 'R_SEP_N210', 'R_SEP_N215', 'R_SEP_N220', 'R_SEP_N230', 'R_SEP_N350', 'R_SEP_N410', 'R_SEP_N450'],
                'SXV': ['R_SXV_N110', 'R_SXV_N180', 'R_SXV_N210', 'R_SXV_N220', 'R_SXV_N230', 'R_SXV_N240', 'R_SXV_N350'],
                'SYS': ['R_SYS_N100', 'R_SYS_N120', 'R_SYS_N180', 'R_SYS_N210', 'R_SYS_N220', 'R_SYS_N230', 'R_SYS_N250', 'R_SYS_N260', 'R_SYS_N270', 'R_SYS_N310', 'R_SYS_N315', 'R_SYS_N320', 'R_SYS_N325', 'R_SYS_N330', 'R_SYS_N335', 'R_SYS_N337', 'R_SYS_N340', 'R_SYS_N350'],
                'TTQ': ['R_TTQ_N110', 'R_TTQ_N120'],
                'UHF': ['R_UHF_N110', 'R_UHF_N180', 'R_UHF_N350', 'R_UHF_N352'],
                'XTX': ['R_XTX_N110', 'R_XTX_N180', 'R_XTX_N210', 'R_XTX_N220', 'R_XTX_N230', 'R_XTX_N240', 'R_XTX_N350', 'R_XTX_N610']
               }
    TTQ_contingency = {'ADC': ['TT_ADC_C310'],
                'GPS': ['TT_GPS_C120'],
                'UHF': ['TT_UHF_C210'],
               }
    TTQ_nominal = {'ADC': ['TT_ADC_N110', 'TT_ADC_N120', 'TT_ADC_N210', 'TT_ADC_N220', 'TT_ADC_N230', 'TT_ADC_N240', 'TT_ADC_N310', 'TT_ADC_N520'],
                'ADF': ['TT_ADF_N110', 'TT_ADF_N180', 'TT_ADF_N210', 'TT_ADF_N220', 'TT_ADF_N230', 'TT_ADF_N235', 'TT_ADF_N240', 'TT_ADF_N250', 'TT_ADF_N260', 'TT_ADF_N270', 'TT_ADF_N280', 'TT_ADF_N285'],
                'CAM': ['TT_CAM_N110', 'TT_CAM_N180', 'TT_CAM_N210', 'TT_CAM_N220'],
                'CCS': ['TT_CCS_N110', 'TT_CCS_N180', 'TT_CCS_N230', 'TT_CCS_N240', 'TT_CCS_N410', 'TT_CCS_N430'],
                'DHS': ['TT_DHS_N110', 'TT_DHS_N210', 'TT_DHS_N215', 'TT_DHS_N410', 'TT_DHS_N420', 'TT_DHS_N425', 'TT_DHS_N450'],
                'EPS': ['TT_EPS_N110', 'TT_EPS_N112', 'TT_EPS_N120', 'TT_EPS_N122', 'TT_EPS_N125', 'TT_EPS_N127', 'TT_EPS_N130', 'TT_EPS_N132', 'TT_EPS_N135', 'TT_EPS_N137', 'TT_EPS_N140', 'TT_EPS_N150', 'TT_EPS_N160', 'TT_EPS_N180', 'TT_EPS_N510'],
                'EXP': ['TT_EXP_N310', 'TT_EXP_N320'],
                'GPS': ['TT_GPS_N110', 'TT_GPS_N120', 'TT_GPS_N130', 'TT_GPS_N210', 'TT_GPS_N220', 'TT_GPS_N510'],
                'OBC': ['TT_OBC_N110', 'TT_OBC_N180', 'TT_OBC_N420', 'TT_OBC_N425', 'TT_OBC_N522', 'TT_OBC_N527'],
                'ORX': ['TT_ORX_N110', 'TT_ORX_N180', 'TT_ORX_N210', 'TT_ORX_N220'],
                'SDR': ['TT_SDR_N110', 'TT_SDR_N180', 'TT_SDR_N210', 'TT_SDR_N220'],
                'SEP': ['TT_SEP_N110', 'TT_SEP_N180', 'TT_SEP_N210', 'TT_SEP_N215', 'TT_SEP_N220', 'TT_SEP_N230', 'TT_SEP_N450' ],
                'SXV': ['TT_SXV_N110', 'TT_SXV_N180', 'TT_SXV_N210', 'TT_SXV_N220', 'TT_SXV_N230', 'TT_SXV_N240'],
                'SYS': ['TT_SYS_N100', 'TT_SYS_N120', 'TT_SYS_N180', 'TT_SYS_N210', 'TT_SYS_N220', 'TT_SYS_N230', 'TT_SYS_N260', 'TT_SYS_N270', 'TT_SYS_N310', 'TT_SYS_N315', 'TT_SYS_N320', 'TT_SYS_N325', 'TT_SYS_N330', 'TT_SYS_N335', 'TT_SYS_N337', 'TT_SYS_N340'],
                'UHF': ['TT_UHF_N110', 'TT_UHF_N180'],
                'XTX': ['TT_XTX_N110', 'TT_XTX_N180', 'TT_XTX_N210', 'TT_XTX_N220', 'TT_XTX_N230', 'TT_XTX_N240']
               }
    # creating tree structure of repository
    procedures = {'decommissioning_nominal': decommissioning_nominal, 'LEOP_contingency': LEOP_contingency, 'LEOP_nominal': LEOP_nominal, 'Routine_contingency': routine_contingency, 'Routine_nominal': routine_nominal, 'TTQ_contingency': TTQ_contingency, 'TTQ_nominal': TTQ_nominal}
    SSM = {'Procedures': procedures}
    tree_structure_PROCEDURE_repository = {'SSM': SSM}

# VISUALISATION: check if all the Procedures are in the correct folder in the Repository in MATIS (for drag and drop)
    # print all contingency Procedures starts with their corresponding category
    # print('\n\nRoutine_contingency')
    # for params in routine_contingency:
    #     print('\t' + params)
    #     for pr in routine_contingency['{par}'.format(par=params)]:
    #         print('\t\t' + pr)
    #     print()
    # # print all nominal Procedures starts with their corresponding category
    # print('\n\nRoutine_nominal')
    # for params in routine_nominal:
    #     print('\t' + params)
    #     for pr in routine_nominal['{par}'.format(par=params)]:
    #         print('\t\t' + pr)
    #     print()
    return tree_structure_PROCEDURE_repository


def generate_code(ws_front_page, ws_procedure, _BUFFER_FOR_TEXT_FOR_UNKNOWN_COMMANDS, identifier_matrix, tree_structure_params_repository, asdf_list):
    '''
    Generates the PLUTO code and contains the overall logic how
    and when procedures are called.
    Also outputs a text buffer for unknown commands and list that can be used for debugging.
    :param ws_front_page, ws_procedure, _BUFFER_FOR_TEXT_FOR_UNKNOWN_COMMANDS, identifier_matrix, tree_structure_params_repository, asdf_list:
    :return _BUFFER_FOR_TEXT_FOR_UNKNOWN_COMMANDS, asdf_list:
    '''
    f = open(_FILE_NAME, 'a')
    indents = 0
    procedure_title = str(ws_front_page[_PROCEDURE_TITLE_CELL].value).replace(' ', '_').replace('-', '_').replace('\\', '_').replace('/', '_').replace('(', '').replace(')', '').replace(':', '').replace(';', '').replace('+', 'PLUS').replace(',', '').replace('$', '')
    procedure_ID = str(ws_front_page[_PROCEDURE_ID_CELL].value).replace(' ', '_').replace('-', '_').replace('\\', '_').replace('/', '_').replace('(', '').replace(')', '').replace(':', '').replace(';', '').replace('+', 'PLUS').replace(',', '').replace('$', '')
    write_into_f(f, indents, 'procedure\n')
    indents = indent_add(indents)
    write_into_f(f, indents, 'initiate and confirm step {ID}_{TITLE}\n'.format(ID=procedure_ID, TITLE=procedure_title))
    indents = indent_add(indents)
    matrix_indicator1_old = None
    matrix_indicator_operator_old = None
    flag_in_SELECT_CASE = 0
    flag_in_CALL_PROCEDURE = 0
    flag_in_LEN_loop = 0
    matrix_iteration = 1
    last_OPERATION_cell_value = 0
## Start: write "declaration of variables" into global step in PLUTO
    flag_array_TM_CHECK_VARIABLES = [1, 1, 1, 1, 1, 1, 1]
    flag_first_declarable_variable = 1
    array_declared_variables = []
    flag_DECLARE_VARIABLES = 0
    flag_some_variable_declared = 0         # flag to write no comma behind declared variable
    for matrix_line in identifier_matrix[1:]:
        matrix_row_number = matrix_line[0]
        matrix_indicator1 = matrix_line[1]
        try:
            future_matrix_indicator1 = identifier_matrix[matrix_iteration + 1][1]
            future_STEP_cell_value, future_OPERATIONS_cell_value, future_ID_cell_value, future_DESCRIPTION_cell_value, future_TYPE_cell_value, future_RAW_cell_value, future_ENG_cell_value, future_UNIT_cell_value = get_current_row_cells(ws_procedure, identifier_matrix[matrix_iteration + 1][0])
        except:
            print('', end='')
        matrix_indicator_operator = matrix_line[2]
        current_STEP_cell_value, current_OPERATIONS_cell_value, current_ID_cell_value, current_DESCRIPTION_cell_value, current_TYPE_cell_value, current_RAW_cell_value, current_ENG_cell_value, current_UNIT_cell_value = get_current_row_cells(ws_procedure, matrix_row_number)
        # Start DECLARE_VARIABLES section in Excel
        if str(matrix_indicator_operator).replace(' ', '').startswith('DECLAREVARIABLES'):
                flag_DECLARE_VARIABLES = 1
                if flag_first_declarable_variable == 1:
                    flag_first_declarable_variable = 0
                    write_into_f(f, indents, 'declare\n')
                    indents = indent_add(indents)
        elif not 'ID_FIELD' in matrix_indicator1:
            flag_DECLARE_VARIABLES = 0
        if flag_DECLARE_VARIABLES == 1 and str(current_ID_cell_value).replace(' ', '').startswith('$') and not str(current_ID_cell_value).replace(' ', '').replace('$', '').startswith('VAL'):
            indents, array_declared_variables, flag_some_variable_declared = write_DECLARE_VARIABLES(f,
                                                                        future_matrix_indicator1,
                                                                        array_declared_variables,
                                                                        identifier_matrix,
                                                                        matrix_indicator_operator_old,
                                                                        matrix_indicator_operator,
                                                                        current_ID_cell_value,
                                                                        current_DESCRIPTION_cell_value,
                                                                        current_TYPE_cell_value, matrix_iteration,
                                                                        indents, flag_some_variable_declared)
        # End DECLARE_VARIABLES section in Excel
        # Start declaring variables not declared in Excel
        # if str(matrix_indicator_operator).replace(' ', '').startswith(tuple(_KNOWN_OPERATIONS_PARAMETER)) or matrix_indicator1 == 'NEW_OPERATION_STEP':
# TODO: change '>$' to '@$'
        if str(current_TYPE_cell_value) != 'None' and (str(current_RAW_cell_value).replace(' ', '').startswith('@$') or str(current_ENG_cell_value).replace(' ', '').startswith('@$')):
            if flag_first_declarable_variable == 1:
                flag_first_declarable_variable = 0
                write_into_f(f, indents, 'declare\n')
                indents = indent_add(indents)
            indents, flag_array_TM_CHECK_VARIABLES, array_declared_variables, flag_some_variable_declared = write_DECLARE_TM_CHECK_VARIABLES(f,
                                                                                                                array_declared_variables,
                                                                                                                flag_array_TM_CHECK_VARIABLES,
                                                                                                                identifier_matrix,
                                                                                                                matrix_indicator_operator_old,
                                                                                                                matrix_indicator_operator,
                                                                                                                current_ID_cell_value,
                                                                                                                current_DESCRIPTION_cell_value,
                                                                                                                current_TYPE_cell_value,
                                                                                                                matrix_iteration,
                                                                                                                indents, flag_some_variable_declared)
        # End declaring variables not declared in Excel
        matrix_iteration += 1
        matrix_indicator1_old = matrix_indicator1
        matrix_indicator_operator_old = matrix_indicator_operator
    #Remove indents and close variable declaration
    if flag_first_declarable_variable == 0:
        indents = indent_remove(indents)
        write_into_f(f, 0, '\n')
        write_into_f(f, indents, 'end declare\n')
## End: write "declaration of variables" into global step in PLUTO
    matrix_indicator1_old = None
    matrix_indicator_operator_old = None
    flag_in_SELECT_CASE = 0
    flag_in_CALL_PROCEDURE = 0
    flag_in_LEN_loop = 0
    matrix_iteration = 1
    last_OPERATION_cell_value = 0
# Write remaining content
    while matrix_iteration <= len(identifier_matrix[1:]):
        matrix_row_number = identifier_matrix[matrix_iteration][0]
        matrix_indicator1 = identifier_matrix[matrix_iteration][1]
        try:
            future_matrix_indicator1 = identifier_matrix[matrix_iteration+1][1]
            future_STEP_cell_value, future_OPERATIONS_cell_value, future_ID_cell_value, future_DESCRIPTION_cell_value, future_TYPE_cell_value, future_RAW_cell_value, future_ENG_cell_value, future_UNIT_cell_value = get_current_row_cells(ws_procedure, identifier_matrix[matrix_iteration+1][0])
        except:
            print('', end='')
        matrix_indicator_operator = identifier_matrix[matrix_iteration][2]
        current_STEP_cell_value, current_OPERATIONS_cell_value, current_ID_cell_value, current_DESCRIPTION_cell_value, current_TYPE_cell_value, current_RAW_cell_value, current_ENG_cell_value, current_UNIT_cell_value = get_current_row_cells(ws_procedure, matrix_row_number)
        if str(matrix_indicator_operator).replace(' ', '').startswith(tuple(_KNOWN_OPERATIONS_PARAMETER)) or matrix_indicator1 == 'NEW_OPERATION_STEP':
            if str(matrix_indicator_operator).replace(' ', '').startswith('SENDTIMETAG') and not str(matrix_indicator_operator).replace(' ', '').startswith('SENDTIMETAG_'):
                # write into file
                indents = write_SEND(f, ws_procedure, identifier_matrix, current_ID_cell_value,
                                     tree_structure_params_repository,
                                     matrix_indicator_operator, matrix_iteration, current_DESCRIPTION_cell_value,
                                     indents)
                #TODO: ADD functionality for TC check of the command "INSERT OPERATION" -> see function write_CHECK_TCV()
            elif str(matrix_indicator_operator).replace(' ', '').startswith('SENDTIMETAG_'):
                # write into file
                write_SEND_(f, identifier_matrix, matrix_indicator_operator_old, matrix_indicator_operator,
                            current_ID_cell_value,
                            current_TYPE_cell_value, current_DESCRIPTION_cell_value,
                            current_RAW_cell_value, current_ENG_cell_value, matrix_iteration, indents, ws_procedure)
                #TODO: ADD functionality for TC check of the command "INSERT OPERATION" -> see function write_CHECK_TCV()
            elif str(matrix_indicator_operator).replace(' ', '').startswith('SENDANDCHECKTCV') and not str(matrix_indicator_operator).replace(' ', '').startswith('SENDANDCHECKTCV_'):
                # write into file
                indents, flag_in_LEN_loop, last_row_LEN_loop = write_SEND_WITH_TCV(f, ws_procedure, identifier_matrix, current_ID_cell_value, tree_structure_params_repository,
                           matrix_indicator_operator, matrix_iteration, current_DESCRIPTION_cell_value, current_TYPE_cell_value, indents)
            elif not flag_in_LEN_loop and (str(matrix_indicator_operator).replace(' ', '').startswith('SENDANDCHECKTCV_') or (str(matrix_indicator_operator).replace(' ', '').startswith('SENDANDCHECKTCV') and str(current_OPERATIONS_cell_value).replace(' ','').startswith('LEN'))):
                # write into file
                indents = write_SEND_WITH_TCV_(f, identifier_matrix, matrix_indicator_operator_old, matrix_indicator_operator, current_ID_cell_value,
                            current_TYPE_cell_value, current_DESCRIPTION_cell_value,
                            current_RAW_cell_value, current_ENG_cell_value, matrix_iteration, indents, ws_procedure)
            elif str(matrix_indicator_operator).replace(' ', '').startswith('SEND') and not str(matrix_indicator_operator).replace(' ', '').startswith('SEND_') and not str(matrix_indicator_operator).replace(' ', '').startswith('SENDANDCHECKTCV'):
                # write into file
                indents = write_SEND(f, ws_procedure, identifier_matrix, current_ID_cell_value, tree_structure_params_repository,
                           matrix_indicator_operator, matrix_iteration, current_DESCRIPTION_cell_value, indents)
            elif not flag_in_LEN_loop and str(matrix_indicator_operator).replace(' ', '') == ('SEND_'):
                # write into file
                indents = write_SEND_(f, identifier_matrix, matrix_indicator_operator_old, matrix_indicator_operator, current_ID_cell_value,
                            current_TYPE_cell_value, current_DESCRIPTION_cell_value,
                            current_RAW_cell_value, current_ENG_cell_value, matrix_iteration, indents, ws_procedure)
#TODO: if ENG values are added, add functionality of differentiation for RAW and ENG values
            elif str(matrix_indicator_operator).replace(' ', '').startswith('CHECKTM'):
                TC_commands_or_TM_params_starting_category, TC_TM_categories, MIB_TCs_or_TMs, TC_and_TM, SSM, current_ID_cell_value = check_if_TC_or_TM_ID_applicable_and_give_dependencies_in_repository_in_MATIS(
                    current_ID_cell_value, tree_structure_params_repository)
                indents = write_CHECKTM(f, array_declared_variables, TC_TM_categories, MIB_TCs_or_TMs, TC_and_TM, SSM,
                              current_RAW_cell_value, current_ENG_cell_value, current_ID_cell_value,
                              current_DESCRIPTION_cell_value, current_TYPE_cell_value, indents)
            elif str(matrix_indicator_operator).replace(' ', '').startswith('DECLAREVARIABLES'):
                print('', end='') #do nothing (has already been done in upper for loop for global variables. But if removed it will appear in the console as if the command was not found)
            elif matrix_indicator1 == 'NEW_OPERATION_STEP':
                #print(matrix_line)
                #print(current_OPERATIONS_cell_value)
                indents = write_OPERATIONS_and_add_step(f, identifier_matrix, matrix_iteration, current_STEP_cell_value, last_OPERATION_cell_value, current_OPERATIONS_cell_value, indents)
                last_OPERATION_cell_value = current_OPERATIONS_cell_value
            # start of functions for case switch
            elif str(matrix_indicator_operator).replace(' ', '').startswith('SELECTCASE'):
                indents = write_SELECT_CASE(f, matrix_indicator_operator, indents)
                flag_in_SELECT_CASE = 1
            elif str(matrix_indicator_operator).replace(' ', '').startswith('CASE:'):
                indents = write_CASE(f, matrix_indicator_operator, matrix_indicator_operator_old, indents)
            elif str(matrix_indicator_operator).replace(' ', '').startswith('$'):
                indents = write_DOLLAR(f, matrix_indicator_operator_old, matrix_indicator_operator, indents)
            elif str(matrix_indicator_operator).replace(' ', '').startswith('CASEELSE:'):
                indents = write_CASE_ELSE(f, indents)
            elif str(matrix_indicator_operator).replace(' ', '').startswith('ENDCASE'):
                indents = write_END_CASE(f, indents)
                flag_in_SELECT_CASE = 0
            # elif str(matrix_indicator_operator).replace(' ', '').startswith('IF') and str(matrix_indicator_operator).replace(' ', '').endswith('_'):
            #     indents = write_IF_(f, indents, current_DESCRIPTION_cell_value)
            elif str(matrix_indicator_operator).replace(' ', '').startswith('IF') and str(matrix_indicator_operator).replace(' ', '').endswith('THEN'):
                indents = write_IF(f, indents, current_OPERATIONS_cell_value)
            elif str(matrix_indicator_operator).replace(' ', '').startswith('IF') and not str(matrix_indicator_operator).replace(' ', '').endswith('THEN'):
                indents, matrix_iteration = write_IFIN(f, indents, current_OPERATIONS_cell_value, current_ID_cell_value, matrix_iteration, identifier_matrix, ws_procedure, future_matrix_indicator1)
            elif str(matrix_indicator_operator).replace(' ', '').startswith('THEN') and not str(matrix_indicator_operator).replace(' ', '').endswith('RETURN'):
                indents = write_THEN(f, indents, current_OPERATIONS_cell_value)
            elif str(matrix_indicator_operator).replace(' ', '').startswith('ELSEIF'):
                indents, matrix_iteration = write_ELSEIF(f, indents, current_OPERATIONS_cell_value, current_ID_cell_value, matrix_iteration, identifier_matrix, ws_procedure, future_matrix_indicator1)
            elif str(matrix_indicator_operator).replace(' ', '').startswith('ELSE'):
                intents = write_ELSE(f, indents)
            elif str(matrix_indicator_operator).replace(' ', '').startswith('ENDIF'):
                indents = write_END_IF(f, indents)
            elif str(current_OPERATIONS_cell_value).replace(' ', '').startswith('CALLPROCEDURE'):
                indents = write_PROCEDURE(f, indents, identifier_matrix, matrix_iteration)
                flag_in_CALL_PROCEDURE = 1
            elif flag_in_CALL_PROCEDURE == 1 and str(current_OPERATIONS_cell_value).replace(' ', '').startswith('THENRETURN'):
                flag_in_CALL_PROCEDURE = 0
            elif str(matrix_indicator_operator).replace(' ', '').startswith('EXECUTEINTERMINALONMCSMACHINE'):
                indents = write_EXECUTEINTERMINALONMCSMACHINE(f, indents, current_DESCRIPTION_cell_value)
            elif str(matrix_indicator_operator).replace(' ', '').startswith('CALLENGINEER'):
                indents = write_CALLENGINEER(f, indents, matrix_iteration, identifier_matrix)
            elif str(matrix_indicator_operator).replace(' ', '').startswith('WAIT'):
                indents = write_WAIT(f, indents, current_OPERATIONS_cell_value)
#TODO: check if still needed
            elif flag_in_LEN_loop:
                if last_row_LEN_loop == identifier_matrix[matrix_iteration][0]:
                    flag_in_LEN_loop = 0
            else:
                print('STEP: ' + str(current_STEP_cell_value) + '; matrix_indicator_operator: ' + matrix_indicator_operator + '; OPERATION: ' +str(current_OPERATIONS_cell_value) + '; ID: ' + str(current_ID_cell_value))
            # end of functions for case switch
        elif flag_in_CALL_PROCEDURE == 1:
            # ignore everything inside 'CALL PROCEDURE'
            print('', end='')  # check if continue can be used without any bugs
        else:
            # write unknown stuff as a comment into file
            _BUFFER_FOR_TEXT_FOR_UNKNOWN_COMMANDS = write_else(f, identifier_matrix, current_STEP_cell_value, current_OPERATIONS_cell_value, matrix_iteration, current_ID_cell_value,
               current_DESCRIPTION_cell_value, current_TYPE_cell_value, current_RAW_cell_value,
               current_ENG_cell_value, current_UNIT_cell_value, matrix_indicator1, _BUFFER_FOR_TEXT_FOR_UNKNOWN_COMMANDS)
        matrix_iteration += 1
        matrix_indicator1_old = matrix_indicator1
        matrix_indicator_operator_old = matrix_indicator_operator
    indents = indent_remove(indents)
    write_into_f(f, indents, 'end step;\n')
    indents = indent_remove(indents)
    write_into_f(f, indents, 'end procedure\n')
    return _BUFFER_FOR_TEXT_FOR_UNKNOWN_COMMANDS, asdf_list


###############################################
## WRITE FUNCTIONS
###############################################

def write_SEND_WITH_TCV(f, ws_procedure, identifier_matrix, current_ID_cell_value, tree_structure_params_repository, matrix_indicator_operator, matrix_iteration, current_DESCRIPTION_cell_value, current_TYPE_cell_value, indents):
    '''
    Converts and writes the SEND WITH TCV command into output file.
    :param f, ws_procedure, identifier_matrix, current_ID_cell_value, tree_structure_params_repository, matrix_indicator_operator, matrix_iteration, current_DESCRIPTION_cell_value, current_TYPE_cell_value, indents:
    :return indents, flag_in_LEN_loop, last_row_LEN_loop:
    '''
    matrix_iteration_ = matrix_iteration
    matrix_iteration_ += 1
    flag_in_LEN_loop = 0
    last_row_LEN_loop = 0
    TC_commands_or_TM_params_starting_category, TC_TM_categories, MIB_TCs_or_TMs, TC_and_TM, SSM, current_ID_cell_value = check_if_TC_or_TM_ID_applicable_and_give_dependencies_in_repository_in_MATIS(
        current_ID_cell_value, tree_structure_params_repository)
    write_into_f(f, indents, 'initiate and confirm ' + str(current_ID_cell_value) + ' of ' + TC_TM_categories + ' of ' + MIB_TCs_or_TMs + ' of ' + TC_and_TM + ' of ' + SSM)
    if identifier_matrix[matrix_iteration + 1][2] != (matrix_indicator_operator + '_'):
        # support for "with directives" of TC commands (same was done for write_SEND_WITH_TCV_())
        row_number = int(identifier_matrix[matrix_iteration][0]) + 1
        _, _, iteration_ID_cell_value, iteration_DESCRIPTION_cell_value, _, _, _, _ = get_current_row_cells(ws_procedure, row_number)
        if iteration_DESCRIPTION_cell_value != None:
            write_into_f(f, 0, '\t\t\t//{DESCRIPTION_cell}\n'.format(DESCRIPTION_cell=current_DESCRIPTION_cell_value))
            indents = indent_add(indents)
            write_into_f(f, indents, "with directives\n")
            indents = indent_add(indents)
            while iteration_ID_cell_value == None and iteration_DESCRIPTION_cell_value != None:
                _, _, _, iteration_DESCRIPTION_cell_value, _, iteration_RAW_cell_value, _, _ = get_current_row_cells(ws_procedure, row_number)
                with_directives_string = get_with_directives_string(iteration_DESCRIPTION_cell_value, iteration_RAW_cell_value)
                write_into_f(f, indents, with_directives_string)
                row_number += 1
                _, _, _, iteration_DESCRIPTION_cell_value, _, iteration_RAW_cell_value, _, _ = get_current_row_cells(ws_procedure, row_number)
                if iteration_ID_cell_value == None and iteration_DESCRIPTION_cell_value != None:
                    write_into_f(f, 0, ",\n")
                else:
                    write_into_f(f, 0, "\n")
                    indents = indent_remove(indents)
                    write_into_f(f, indents, 'end with;\n')
                    indents = indent_remove(indents)
        else:
            write_into_f(f, 0, ';\t\t\t//{DESCRIPTION_cell}\n'.format(DESCRIPTION_cell=current_DESCRIPTION_cell_value))
    else:
        write_into_f(f, 0, '\t\t\t//{DESCRIPTION_cell}\n'.format(DESCRIPTION_cell=current_DESCRIPTION_cell_value))
        indents = indent_add(indents)
    return indents, flag_in_LEN_loop, last_row_LEN_loop


def get_with_directives_string(DESCRIPTION_cell_value, RAW_cell_value):
    '''
    returns the with directive options string
    :param DESCRIPTION_cell_value, RAW_cell_value:
    :return with_directives_string:
    '''
    # conditions for different cases
    if str(DESCRIPTION_cell_value).replace(' ', '').startswith("DYNAMICPTVOVERRIDE"):
        with_directives_string = "dynamic_ptv := \"overridden\""
    elif str(DESCRIPTION_cell_value).replace(' ', '').startswith("STATICPTVOVERRIDE"):
        with_directives_string = "static_ptv := \"overridden\""
    elif str(DESCRIPTION_cell_value).replace(' ', '').startswith("EXECUTIONTIME"):
        with_directives_string = "execution_time := " + str(RAW_cell_value).replace('$', '').replace('NOW', 'current_time()')
    elif str(DESCRIPTION_cell_value).replace(' ', '').startswith("RELEASETIME"):
        with_directives_string = "release_time := " + str(RAW_cell_value).replace('$', '').replace('NOW', 'current_time()')
    elif str(DESCRIPTION_cell_value).replace(' ', '').startswith("CEVDISABLE"):
        with_directives_string = "execution_verification := \"disabled\""
    else:
        with_directives_string = "//COMMAND HAS NOT YET BEEN DEFINED; Original = " + str(DESCRIPTION_cell_value)
    return with_directives_string


def write_CHECK_TCV():
    #TODO: implement TCV check (see TT-CCS-N110)
    pass


def write_SEND_WITH_TCV_(f, identifier_matrix, matrix_indicator_operator_old, matrix_indicator_operator, current_ID_cell_value, current_TYPE_cell_value, current_DESCRIPTION_cell_value, current_RAW_cell_value, current_ENG_cell_value, matrix_iteration, indents, ws_procedure):
    '''
   Converts and writes the SEND WITH TCV command's further inherent options/text
   into the output file.
   :param f, identifier_matrix, matrix_indicator_operator_old, matrix_indicator_operator, current_ID_cell_value, current_TYPE_cell_value, current_DESCRIPTION_cell_value, current_RAW_cell_value, current_ENG_cell_value, matrix_iteration, indents, ws_procedure:
   :return indents:
   '''
    current_ID_cell_value = check_if_ID_starts_with_digit(current_ID_cell_value)
    if matrix_indicator_operator_old == matrix_indicator_operator:
        indents = write_value_plus_add_type_and_description_as_comment(f, identifier_matrix, matrix_iteration, matrix_indicator_operator, current_TYPE_cell_value, current_DESCRIPTION_cell_value,
                                                             current_RAW_cell_value, current_ENG_cell_value, current_ID_cell_value, indents)
    else:
        # TODO write eng as default for parameter but if raw value existent, write it as comment behind the parameter and behind the description of that specific parameter
        write_into_f(f, indents, 'with arguments\n')
        indents = indent_add(indents)
        indents = write_value_plus_add_type_and_description_as_comment(f, identifier_matrix, matrix_iteration, matrix_indicator_operator, current_TYPE_cell_value, current_DESCRIPTION_cell_value,
                                                             current_RAW_cell_value, current_ENG_cell_value, current_ID_cell_value, indents)
    if identifier_matrix[matrix_iteration + 1][2] != matrix_indicator_operator:
        indents = indent_remove(indents)
        write_into_f(f, indents, 'end with')
        indents = indent_remove(indents)
        row_number = int(identifier_matrix[matrix_iteration][0]) + 1
        _, _, iteration_ID_cell_value, iteration_DESCRIPTION_cell_value, _, _, _, _ = get_current_row_cells(
            ws_procedure, row_number)
        if iteration_DESCRIPTION_cell_value != None:
            write_into_f(f, 0, "\n")
            indents = indent_add(indents)
            write_into_f(f, indents, "with directives\n")
            indents = indent_add(indents)
            while iteration_ID_cell_value == None and iteration_DESCRIPTION_cell_value != None:
                _, _, _, iteration_DESCRIPTION_cell_value, _, iteration_RAW_cell_value, _, _ = get_current_row_cells(
                    ws_procedure, row_number)
                with_directives_string = get_with_directives_string(iteration_DESCRIPTION_cell_value,
                                                                    iteration_RAW_cell_value)
                write_into_f(f, indents, with_directives_string)
                row_number += 1
                _, _, _, iteration_DESCRIPTION_cell_value, _, iteration_RAW_cell_value, _, _ = get_current_row_cells(
                    ws_procedure, row_number)
                if iteration_ID_cell_value == None and iteration_DESCRIPTION_cell_value != None:
                    write_into_f(f, 0, ",\n")
                else:
                    write_into_f(f, 0, "\n")
                    indents = indent_remove(indents)
                    write_into_f(f, indents, 'end with;\n')
        else:
            write_into_f(f, 0, ';\n')
        indents = indent_remove(indents)
    return indents


def write_SEND(f, ws_procedure, identifier_matrix, current_ID_cell_value, tree_structure_params_repository, matrix_indicator_operator, matrix_iteration, current_DESCRIPTION_cell_value, indents):
    '''
    Converts and writes the SEND command into the output file.
    :param f, ws_procedure, identifier_matrix, current_ID_cell_value, tree_structure_params_repository, matrix_indicator_operator, matrix_iteration, current_DESCRIPTION_cell_value, indents:
    :return indents:
    '''
    TC_commands_or_TM_params_starting_category, TC_TM_categories, MIB_TCs_or_TMs, TC_and_TM, SSM, current_ID_cell_value = check_if_TC_or_TM_ID_applicable_and_give_dependencies_in_repository_in_MATIS(current_ID_cell_value, tree_structure_params_repository)
    write_into_f(f, indents, 'initiate ' + str(current_ID_cell_value) + ' of ' + TC_TM_categories + ' of ' + MIB_TCs_or_TMs + ' of ' + TC_and_TM + ' of ' + SSM)
    if identifier_matrix[matrix_iteration + 1][2] != (matrix_indicator_operator + '_'):
        # support for "with directives" of TC commands (same was done for write_SEND_WITH_TCV_())
        row_number = int(identifier_matrix[matrix_iteration][0]) + 1
        _, _, iteration_ID_cell_value, iteration_DESCRIPTION_cell_value, _, _, _, _ = get_current_row_cells(ws_procedure, row_number)
        if iteration_DESCRIPTION_cell_value != None:
            write_into_f(f, 0, '\t\t\t//{DESCRIPTION_cell}\n'.format(DESCRIPTION_cell=current_DESCRIPTION_cell_value))
            indents = indent_add(indents)
            write_into_f(f, indents, "with directives\n")
            indents = indent_add(indents)
            while iteration_ID_cell_value == None and iteration_DESCRIPTION_cell_value != None:
                _, _, _, iteration_DESCRIPTION_cell_value, _, iteration_RAW_cell_value, _, _ = get_current_row_cells(
                    ws_procedure, row_number)
                with_directives_string = get_with_directives_string(iteration_DESCRIPTION_cell_value,
                                                                    iteration_RAW_cell_value)
                write_into_f(f, indents, with_directives_string)
                row_number += 1
                _, _, _, iteration_DESCRIPTION_cell_value, _, iteration_RAW_cell_value, _, _ = get_current_row_cells(
                    ws_procedure, row_number)
                if iteration_ID_cell_value == None and iteration_DESCRIPTION_cell_value != None:
                    write_into_f(f, 0, ",\n")
                else:
                    write_into_f(f, 0, "\n")
                    indents = indent_remove(indents)
                    write_into_f(f, indents, 'end with;\n')
                    indents = indent_remove(indents)
        else:
            write_into_f(f, 0, ';')
            write_into_f(f, 0, '\t\t\t//{DESCRIPTION_cell}\n'.format(DESCRIPTION_cell=current_DESCRIPTION_cell_value))
        write_into_f(f, indents, 'wait for 0.5s;\n')
    else:
        write_into_f(f, 0, '')
        indents = indent_add(indents)
        write_into_f(f, 0, '\t\t\t//{DESCRIPTION_cell}\n'.format(DESCRIPTION_cell=current_DESCRIPTION_cell_value))
    return indents


def write_SEND_(f, identifier_matrix, matrix_indicator_operator_old, matrix_indicator_operator, current_ID_cell_value, current_TYPE_cell_value, current_DESCRIPTION_cell_value, current_RAW_cell_value, current_ENG_cell_value, matrix_iteration, indents, ws_procedure):
    '''
    Converts and writes the SEND command's further inherent options/text
    into the output file.
    :param f, identifier_matrix, matrix_indicator_operator_old, matrix_indicator_operator, current_ID_cell_value, current_TYPE_cell_value, current_DESCRIPTION_cell_value, current_RAW_cell_value, current_ENG_cell_value, matrix_iteration, indents, ws_procedure:
    :return indents:
    '''
    current_ID_cell_value = check_if_ID_starts_with_digit(current_ID_cell_value)
    if matrix_indicator_operator_old == matrix_indicator_operator:
        indents = write_value_plus_add_type_and_description_as_comment(f, identifier_matrix, matrix_iteration, matrix_indicator_operator, current_TYPE_cell_value, current_DESCRIPTION_cell_value,
                                                             current_RAW_cell_value, current_ENG_cell_value, current_ID_cell_value, indents)
    else:
        # TODO write eng as default for parameter but if raw value existent, write it as comment behind the parameter and behind the description of that specific parameter
        write_into_f(f, indents, 'with arguments\n')
        indents = indent_add(indents)
        indents = write_value_plus_add_type_and_description_as_comment(f, identifier_matrix, matrix_iteration, matrix_indicator_operator, current_TYPE_cell_value, current_DESCRIPTION_cell_value,
                                                             current_RAW_cell_value, current_ENG_cell_value, current_ID_cell_value, indents)
    if identifier_matrix[matrix_iteration + 1][2] != matrix_indicator_operator:
        indents = indent_remove(indents)
        write_into_f(f, indents, 'end with\n')
        indents = indent_remove(indents)
        # support for "with directives" of TC commands (same was done for write_SEND_WITH_TCV_())
        row_number = int(identifier_matrix[matrix_iteration][0]) + 1
        _, _, iteration_ID_cell_value, iteration_DESCRIPTION_cell_value, _, _, _, _ = get_current_row_cells(
            ws_procedure, row_number)
        if iteration_DESCRIPTION_cell_value != None:
            indents = indent_add(indents)
            write_into_f(f, indents, "with directives\n")
            indents = indent_add(indents)
            while iteration_ID_cell_value == None and iteration_DESCRIPTION_cell_value != None:
                _, _, _, iteration_DESCRIPTION_cell_value, _, iteration_RAW_cell_value, _, _ = get_current_row_cells(
                    ws_procedure, row_number)
                with_directives_string = get_with_directives_string(iteration_DESCRIPTION_cell_value,
                                                                    iteration_RAW_cell_value)
                write_into_f(f, indents, with_directives_string)
                row_number += 1
                _, _, _, iteration_DESCRIPTION_cell_value, _, iteration_RAW_cell_value, _, _ = get_current_row_cells(
                    ws_procedure, row_number)
                if iteration_ID_cell_value == None and iteration_DESCRIPTION_cell_value != None:
                    write_into_f(f, 0, ",\n")
                else:
                    write_into_f(f, 0, "\n")
                    indents = indent_remove(indents)
                    write_into_f(f, indents, 'end with;\n')
                    indents = indent_remove(indents)
        else:
            write_into_f(f, indents, 'end with;\n')
            indents = indent_remove(indents)
        write_into_f(f, indents, 'wait for 0.5s;\n')
    return indents


def write_OPERATIONS_and_add_step(f, identifier_matrix, matrix_iteration, current_STEP_cell_value, last_OPERATION_cell_value, current_OPERATIONS_cell_value, indents):
    '''
    Writes operations and adds step
    :param f, identifier_matrix, matrix_iteration, current_STEP_cell_value, last_OPERATION_cell_value, current_OPERATIONS_cell_value, indents:
    :return indents:
    '''
    if last_OPERATION_cell_value != 0:
        indents = indent_remove(indents)
        write_into_f(f, indents, 'end step;\n')
    write_into_f(f, indents, '\n// STEP: {STEP}, OPERATION: {OPERATION}\n'.format(STEP=current_STEP_cell_value, OPERATION=current_OPERATIONS_cell_value))
    if current_OPERATIONS_cell_value != identifier_matrix[-1][-1]:
        write_into_f(f, indents, 'initiate and confirm step ' + str(current_OPERATIONS_cell_value).replace(' ', '_').replace('-', '_').replace('\\', '_').replace('/', '_').replace('(', '').replace(')', '').replace(':', '').replace(';', '').replace('+', 'PLUS').replace(',', '').replace('$', '') + '\n')
        indents = indent_add(indents)
    return indents


def write_else(f, identifier_matrix, current_STEP_cell_value, current_OPERATIONS_cell_value, matrix_iteration, current_ID_cell_value, current_DESCRIPTION_cell_value, current_TYPE_cell_value, current_RAW_cell_value, current_ENG_cell_value, current_UNIT_cell_value, matrix_indicator1, _BUFFER_FOR_TEXT_FOR_UNKNOWN_COMMANDS):
    '''
    Writes else command into file.
    :param f, identifier_matrix, current_STEP_cell_value, current_OPERATIONS_cell_value, matrix_iteration, current_ID_cell_value, current_DESCRIPTION_cell_value, current_TYPE_cell_value, current_RAW_cell_value, current_ENG_cell_value, current_UNIT_cell_value, matrix_indicator1, _BUFFER_FOR_TEXT_FOR_UNKNOWN_COMMANDS:
    :return _BUFFER_FOR_TEXT_FOR_UNKNOWN_COMMANDS:
    '''
    mini_row_content_buffer = '// STEP: {STEP}; OPERATION: {OPERATION}; ID: {ID}; DESCRIPTION: {DESCRIPTION}; TYPE: {TYPE}; RAW: {RAW}; ENG: {ENG}; UNIT: {UNIT}\n'.format(
        STEP=current_STEP_cell_value, OPERATION=current_OPERATIONS_cell_value, ID=current_ID_cell_value,
        DESCRIPTION=current_DESCRIPTION_cell_value, TYPE=current_TYPE_cell_value, RAW=current_RAW_cell_value,
        ENG=current_ENG_cell_value, UNIT=current_UNIT_cell_value)
    _BUFFER_FOR_TEXT_FOR_UNKNOWN_COMMANDS += mini_row_content_buffer.replace('\n', ' ') + '\n'
    if str(identifier_matrix[matrix_iteration + 1][2]).replace(' ', '').startswith(tuple(_KNOWN_OPERATIONS_PARAMETER)) or matrix_indicator1 == 'NEW_OPERATION_STEP':
        f.write('\n//////////////////////////////////////\n')
        f.write('// TODO: UNIDENTIFIED COMMENT(S)\n')
        f.write(_BUFFER_FOR_TEXT_FOR_UNKNOWN_COMMANDS)
        f.write('// END UNIDENTIFIED COMMENT(S)\n')
        f.write('//////////////////////////////////////\n\n')
        _BUFFER_FOR_TEXT_FOR_UNKNOWN_COMMANDS = ''
    return _BUFFER_FOR_TEXT_FOR_UNKNOWN_COMMANDS


def write_DECLARE_VARIABLES(f, future_matrix_indicator1, array_declared_variables, identifier_matrix, matrix_indicator_operator_old, matrix_indicator_operator, current_ID_cell_value, current_DESCRIPTION_cell_value, current_TYPE_cell_value, matrix_iteration, indents, flag_some_variable_declared):
    '''
    Writes Declare variables command into file.
    :param f, future_matrix_indicator1, array_declared_variables, identifier_matrix, matrix_indicator_operator_old, matrix_indicator_operator, current_ID_cell_value, current_DESCRIPTION_cell_value, current_TYPE_cell_value, matrix_iteration, indents, flag_some_variable_declared:
    :return indents, array_declared_variables, flag_some_variable_declared:
    '''
    variable_without_dollar_sign = str(current_ID_cell_value).replace('$', '').replace(' ', '')
    array_declared_variables.append(variable_without_dollar_sign)
    # if matrix_indicator_operator == 'DECLARE VARIABLES_':    #matrix_indicator_operator or matrix_indicator_operator_old+'_' == matrix_indicator_operator:
    write_into_f(f, indents, 'variable {VARIABLE} of type {TYPE} described by \"{DESCRIPTION}\"'.format(VARIABLE=variable_without_dollar_sign, TYPE=current_TYPE_cell_value, DESCRIPTION=current_DESCRIPTION_cell_value))
    if future_matrix_indicator1 != 'NEW_OPERATION_FIELD':
        write_into_f(f, 0, ',\n')
    else:
        write_into_f(f, 0, '')
    flag_some_variable_declared = 1
    return indents, array_declared_variables, flag_some_variable_declared


def write_SELECT_CASE(f, matrix_indicator_operator, indents):
    '''
    Writes SELECT CASE command into file.
    :param f, matrix_indicator_operator, indents:
    :return indents:
    '''
    #print(matrix_indicator_operator)
    try:
        parameter = str(re.findall(r'\$\w+', matrix_indicator_operator)[0]).replace('$', '')
        write_into_f(f, indents, 'in case {PARAMETER}\n'.format(PARAMETER=parameter))
    except IndexError:
            write_into_f(f, indents, '//NON STANDARD COMMAND: ' + matrix_indicator_operator)
    indents = indent_add(indents)
    return indents


def write_CASE(f, matrix_indicator_operator, matrix_indicator_operator_old, indents):
    '''
    Writes CASE command into file.
    :param f, matrix_indicator_operator, matrix_indicator_operator_old, indents:
    :return indents:
    '''
    if str(matrix_indicator_operator_old).replace(' ', '').startswith('SELECT'):
        parameter = str(matrix_indicator_operator).split(':', 1)[1].replace(' ', '')
        write_into_f(f, indents, 'is = {PARAMETER}:\n'.format(PARAMETER=parameter.replace('\n', '')))
        indents = indent_add(indents)
    else:
        indents = indent_remove(indents)
        parameter = str(matrix_indicator_operator).split(':', 1)[1].replace(' ', '')
        write_into_f(f, indents, 'or is = {PARAMETER}:\n'.format(PARAMETER=parameter.replace('\n', '')))
        indents = indent_add(indents)
    return indents


def write_CASE_ELSE(f, indents):
    '''
    Writes CASE ELSE command into file.
    :param f, indents:
    :return indents:
    '''
    indents = indent_remove(indents)
    write_into_f(f, indents, 'otherwise:\n')
    indents = indent_add(indents)
    return indents


def write_END_CASE(f, indents):
    '''
    Writes END ELSE command into file.
    :param f, indents:
    :return indents:
    '''
    indents = indent_remove(indents)
    indents = indent_remove(indents)
    write_into_f(f, indents, 'end case;\n')
    return indents


def write_IF(f, indents, current_OPERATIONS_cell_value):
    '''
    Writes IF command into file.
    :param f, indents, current_OPERATIONS_cell_value:
    :return indents:
    '''
    if_condition_content = str(current_OPERATIONS_cell_value).replace('$', '')
    try:
        if_condition_content = if_condition_content.split('IF ', 1)[1].rsplit(' THEN', 1)[0]
        write_into_f(f, indents, 'if ' + str(if_condition_content).replace('==', '=') + ' then\n')
        indents = indent_add(indents)
    except:
        f.write('\n//////////////////////////////////////\n')
        f.write('// TODO: UNIDENTIFIED COMMENT(S)\n')
        f.write('// WARNING: Could not resolve IF loop.\n // OPERATION: ' + if_condition_content)
        f.write('// END UNIDENTIFIED COMMENT(S)\n')
        f.write('//////////////////////////////////////\n\n')
    return indents


def write_IFIN(f, indents, current_OPERATIONS_cell_value, current_ID_cell_value, matrix_iteration, identifier_matrix, ws_procedure, future_matrix_indicator1):
    '''
    Writes IFIN command into file.
    :param f, indents, current_OPERATIONS_cell_value, current_ID_cell_value, matrix_iteration, identifier_matrix, ws_procedure, future_matrix_indicator1:
    :return indents, matrix_iteration:
    '''
# This function can only be used for comparison of the parameter given and the ID field -> procedure == 'compare if ID of param is the same as ID field'
    if_condition_content = str(current_OPERATIONS_cell_value).replace('$', '').split('IF ', 1)[1].rsplit(' IN', 1)[0]
    write_into_f(f, indents, 'if ')
    if current_ID_cell_value == None:
        matrix_iteration += 1
    while future_matrix_indicator1 == 'NEW_ID_FIELD' or future_matrix_indicator1 == 'FOLLOW_ID_FIELD':
        matrix_line = identifier_matrix[matrix_iteration]
        matrix_row_number = matrix_line[0]
        try:
            future_matrix_indicator1 = identifier_matrix[matrix_iteration + 1][1]
            _, _, future_ID_cell_value, _, _, _, _, _ = get_current_row_cells(ws_procedure, identifier_matrix[matrix_iteration + 1][0])
        except:
            future_ID_cell_value = ''
            print('', end='')
        _, _, current_ID_cell_value, _, _, _, _, _ = get_current_row_cells(ws_procedure, matrix_row_number)
        write_into_f(f, 0, if_condition_content + ' = "' + current_ID_cell_value + '" ')
        if future_matrix_indicator1 == 'NEW_ID_FIELD' or future_matrix_indicator1 == 'FOLLOW_ID_FIELD':
            write_into_f(f, 0, 'or ')
        else:
            write_into_f(f, 0, 'then\n')
            indents = indent_add(indents)
        matrix_iteration += 1
    matrix_line = identifier_matrix[matrix_iteration]
    matrix_row_number = matrix_line[0]
    _, _, current_ID_cell_value, _, _, _, _, _ = get_current_row_cells(ws_procedure, matrix_row_number)
    # write_into_f(f, indents, if_condition_content + ' = "' + current_ID_cell_value + '" then\n')
    matrix_iteration -= 1
    indents = indent_add(indents)
    return indents, matrix_iteration


def write_THEN(f, indents, current_OPERATIONS_cell_value):
    '''
    Writes THEN command into file.
    :param f, indents, current_OPERATIONS_cell_value:
    :return indents:
    '''
    content_to_write = str(current_OPERATIONS_cell_value).split('THEN ', 1)[1].replace('$', '')
    write_into_f(f, indents, content_to_write)
    write_into_f(f, 0, '\n')
    return indents


def write_ELSEIF(f, indents, current_OPERATIONS_cell_value, current_ID_cell_value, matrix_iteration, identifier_matrix, ws_procedure, future_matrix_indicator1):
    '''
    Writes ELSEIF command into file.
    :param f, indents, current_OPERATIONS_cell_value, current_ID_cell_value, matrix_iteration, identifier_matrix, ws_procedure, future_matrix_indicator1:
    :return indents, matrix_iteration:
    '''
    if_condition_content = str(current_OPERATIONS_cell_value).replace('$', '')
    try:
        indents = indent_remove(indents)
        if_condition_content = if_condition_content.split('ELSEIF ', 1)[1].rsplit(' THEN', 1)[0]
        write_into_f(f, indents, 'else if ' + str(if_condition_content).replace('==', '=') + ' then\n')
        indents = indent_add(indents)
    except:
        try:
            #see if PARAMETER is IN certain ID values (else if -> normal if for this case)
            indents = indent_remove(indents)
            if_condition_content = if_condition_content.split('ELSE IF ', 1)[1].rsplit(' IN', 1)[0]
            write_into_f(f, indents, 'end if;\n')
            indents, matrix_iteration = write_IFIN(f, indents, current_OPERATIONS_cell_value, current_ID_cell_value, matrix_iteration, identifier_matrix, ws_procedure, future_matrix_indicator1)
        except:
            f.write('\n//////////////////////////////////////\n')
            f.write('// TODO: UNIDENTIFIED COMMENT(S)\n')
            f.write('// WARNING: Could not resolve IF loop.\n // OPERATION: ' + if_condition_content)
            f.write('// END UNIDENTIFIED COMMENT(S)\n')
            f.write('//////////////////////////////////////\n\n')
    return indents, matrix_iteration


def write_END_IF(f, indents):
    '''
    Writes END IF command into file.
    :param f, indents:
    :return indents:
    '''
    indents = indent_remove(indents)
    write_into_f(f, indents, 'end if;\n')
    return indents


def write_DOLLAR(f, matrix_indicator_operator_old, matrix_indicator_operator, indents):
    '''
    Writes DOLLAR command into file.
    :param f, matrix_indicator_operator_old, matrix_indicator_operator, indents:
    :return indents:
    '''
    case_content = str(matrix_indicator_operator).split('$', 1)[1].replace('$', '').replace('=', ':=').replace(':=:=', ':=').replace('::', ':')
    write_into_f(f, indents, '{CASE_CONTENT};\n'.format(CASE_CONTENT=case_content))
    return indents


def write_PROCEDURE(f, indents, identifier_matrix, matrix_iteration_):
    '''
    Writes PROCEDURE command into file.
    :param f, indents, identifier_matrix, matrix_iteration_:
    :return indents:
    '''
    procedure_PARAMS = []
    matrix_iteration_ += 1
    procedure_ID = identifier_matrix[matrix_iteration_][2]
    matrix_iteration_ += 1
    procedure_TITLE = str(identifier_matrix[matrix_iteration_][2]).split('TITLE:')[1]
    matrix_iteration_ += 1
    procedure_REASON = str(identifier_matrix[matrix_iteration_][2]).split('REASON:')[1]
    matrix_iteration_ += 2
    try:
        while not str(identifier_matrix[matrix_iteration_][2]).replace(' ', '').startswith('THENRETURN'):
            try:
                procedure_PARAMS.append(str(str(identifier_matrix[matrix_iteration_][2]).split('.', 1)[1]))
            except IndexError:
                write_into_f(f, indents, '//NON STANDARD COMMAND: ' + identifier_matrix[matrix_iteration_][2])
            finally:
                matrix_iteration_ += 1
    except IndexError:
        print('An anomaly in corresponding Excel file has been found in: ' + _FILE_NAME)
        print('Likely a \"THEN RETURN\" is missing')
    write_into_f(f, indents, '// CALL PROCEDURE: {ID}\n'.format(ID=procedure_ID))
    write_into_f(f, indents, '// TITLE: {TITLE}\n'.format(TITLE=procedure_TITLE.replace('\n', ' ')))
    write_into_f(f, indents, '// REASON: {REASON}\n'.format(REASON=str(procedure_REASON).replace('\n', ' ').replace('  ', '')))
    procedure_name_with_underscores = str(procedure_ID).split('ID:')[1].replace(' ', '').replace('-', '_')
    tree_structure_PROCEDURE_repository = create_PROCEDURE_dictionary()
    Procedure_ID, Routine_Category, Routines, Procedures, SSM, procedure_name_with_underscores = check_if_PROCEDURE_ID_applicable_and_give_dependencies_in_repository_in_MATIS(procedure_name_with_underscores, tree_structure_PROCEDURE_repository)
    write_into_f(f, indents, 'initiate and confirm {ID_of_procedure} of {ROUTINE_CATEGORY} of {ROUTINES} of {PROCEDURES} of {SSM}'.format(ID_of_procedure=procedure_name_with_underscores, ROUTINE_CATEGORY=Routine_Category, ROUTINES=Routines, PROCEDURES=Procedures, SSM=SSM))

    procedure_PARAMS_length = len(procedure_PARAMS)
    if procedure_PARAMS_length >=1:
        write_into_f(f, 0, '\n')
        indents = indent_add(indents)
        write_into_f(f, indents, 'with arguments\n')
        indents = indent_add(indents)
        for params in procedure_PARAMS:
            procedure_PARAMS_length -= 1
            write_into_f(f, indents, str(params).replace('=', ':=').replace('::', ':').replace('$', ''))
            if procedure_PARAMS_length >= 1:
                write_into_f(f, 0, ',\n')
            else:
                write_into_f(f, 0, '\n')
        indents = indent_remove(indents)
        write_into_f(f, indents, 'end with;\n')
        indents = indent_remove(indents)
    else:
        write_into_f(f, 0, ';\n')
    return indents

#TODO: finish this function and test behavior in MATIS with strings put together
def write_EXECUTEINTERMINALONMCSMACHINE(f, indents, current_DESCRIPTION_cell_value):
    '''
    Writes EXECUTE IN TERMINAL ON MCS MACHINE command into file.
    :param f, indents, current_DESCRIPTION_cell_value:
    :return indents:
    '''
    # SpaceShell
    # if str(current_DESCRIPTION_cell_value).replace(' ', '').startswith('$ground_port_linux_root'):
    #     content_spaceshell = str(current_DESCRIPTION_cell_value).split('"', 1)[1].rsplit('"', 1)[0]
    #     print('asdf' + content_spaceshell)
    content2execute = str(current_DESCRIPTION_cell_value).replace('64 11666', '65 11777').replace('$', '').replace('\"', '\\\"')
    #TODO: replace arguments starting with $ with actual content of procedure, which means "First "+ $VAR + " Second"
    write_into_f(f, indents, 'initiate and confirm execute_and_get_return of Command_Line of SSM\n')
    indents = indent_add(indents)
    write_into_f(f, indents, 'with arguments\n')
    indents = indent_add(indents)
    write_into_f(f, indents, 'COMMAND := \"{CONTENT2EXECUTE}\"\n'.format(CONTENT2EXECUTE=content2execute))
    indents = indent_remove(indents)
    write_into_f(f, indents, 'end with;\n')
    indents = indent_remove(indents)
    return indents


def write_CALLENGINEER(f, indents, matrix_iteration, identifier_matrix):
    '''
    Writes CALL ENGINEER command into file.
    :param f, indents, matrix_iteration, identifier_matrix:
    :return indents:
    '''
    message = identifier_matrix[matrix_iteration+1][2]
    write_into_f(f, indents, 'initiate and confirm Send of Email of Communicator of SwissKnife of PRIME of D0 of TEST_MISSION of SMF\n')
    indents = indent_add(indents)
    write_into_f(f, indents, 'with arguments\n')
    indents = indent_add(indents)
    write_into_f(f, indents, 'Subject := \"MATIS: CALL ENGINEER\",\n')
    write_into_f(f, indents, 'Message := \"{MESSAGE}\",\n'.format(MESSAGE=message.replace('\n', ' ')))
    write_into_f(f, indents, 'ToMail := \"Felix.Hessinger@gmail.com\",\n')
    write_into_f(f, indents, 'ToName := \"Operators\"\n')
    indents = indent_remove(indents)
    write_into_f(f, indents, 'end with;\n')
    indents = indent_remove(indents)
#TODO: abort rest of task/procedure (don't stop calendar or schedule)
    return indents


def write_WAIT(f, indents, current_OPERATIONS_cell_value):
    '''
    Writes WAIT command into file.
    :param f, indents, current_OPERATIONS_cell_value:
    :return indents:
    '''
    try:
        waitingTime = str(current_OPERATIONS_cell_value).split('WAIT FOR ')[1].replace('$', '')
    except:
        waitingTime = "WAITING TIME HAS NOT BEEN FOUND, PLEASE COMPARE WITH EXCEL PROCEDURE"
    write_into_f(f, indents, "wait for " + waitingTime + ";\n")
    return indents


def write_WITH_DIRECTIVES(f, indents, matrix_iteration, identifier_matrix):
    '''
    Ignores WITH DIRECTIVES command since it is already inside of other functions.
    :param f, indents, matrix_iteration, identifier_matrix:
    :return indents:
    '''
    write_into_f(f, indents, "")
    return indents


def write_ELSE(f, indents):
    '''
    Writes ELSE command into file.
    :param f, indents:
    :return indents:
    '''
    indents = indent_remove(indents)
    write_into_f(f, indents, "else\n")
    indents = indent_add(indents)
    return indents

#####################################
## Checks and conversion
#####################################

def check_if_TC_or_TM_ID_applicable_and_give_dependencies_in_repository_in_MATIS(current_ID_cell_value, tree_structure_params_repository):
    '''
    checks if TC or TM id are applicable and outputs the dependencies in the MATIS repository.
    :param current_ID_cell_value, tree_structure_params_repository:
    :return TC_commands_or_TM_params_starting_category, TC_TM_categories, MIB_TCs_or_TMs, TC_and_TM, SSM, current_ID_cell_value_checked_for_ID_starts_with_digit:
    '''
    for SSM in tree_structure_params_repository:
        for TC_and_TM in tree_structure_params_repository[SSM]:
            for MIB_TCs_or_TMs in tree_structure_params_repository[SSM][TC_and_TM]:
                for TC_TM_categories in tree_structure_params_repository[SSM][TC_and_TM][MIB_TCs_or_TMs]:
                    for TC_commands_or_TM_params_starting_category in tree_structure_params_repository[SSM][TC_and_TM][MIB_TCs_or_TMs][TC_TM_categories]:
                        if str(current_ID_cell_value).startswith(TC_commands_or_TM_params_starting_category):
                            # the check_if_ID_starts_with_digit(current_ID_cell_value) is just a safety feature for future procedure names/IDs, if it is decided to let them also start with a digit
                            return TC_commands_or_TM_params_starting_category, TC_TM_categories, MIB_TCs_or_TMs, TC_and_TM, SSM, check_if_ID_starts_with_digit(current_ID_cell_value)
    # the check_if_ID_starts_with_digit(current_ID_cell_value) is just a safety feature for future procedure names/IDs, if it is decided to let them also start with a digit
    return 'SOME_TC_and_TM', 'not inside', 'SOME_TC_and_TM', 'not inside', 'not inside', check_if_ID_starts_with_digit(current_ID_cell_value)


def check_if_PROCEDURE_ID_applicable_and_give_dependencies_in_repository_in_MATIS(procedure_name_with_underscores, tree_structure_PROCEDURE_repository):
    '''
    checks if procedure id are applicable and outputs the dependencies in the MATIS repository.
    :param procedure_name_with_underscores, tree_structure_PROCEDURE_repository:
    :return Procedure_ID, Routine_Category, Routines, Procedures, SSM, check_if_ID_starts_with_digit(procedure_name_with_underscores):
    '''
    for SSM in tree_structure_PROCEDURE_repository:
        for Procedures in tree_structure_PROCEDURE_repository[SSM]:
            for Routines in tree_structure_PROCEDURE_repository[SSM][Procedures]:
                for Routine_Category in tree_structure_PROCEDURE_repository[SSM][Procedures][Routines]:
                    for Procedure_ID in tree_structure_PROCEDURE_repository[SSM][Procedures][Routines][Routine_Category]:
                        if procedure_name_with_underscores.startswith(Procedure_ID):
                            # the check_if_ID_starts_with_digit(current_ID_cell_value) is just a safety feature for future procedure names/IDs, if it is decided to let them also start with a digit
                            return Procedure_ID, Routine_Category, Routines, Procedures, SSM, check_if_ID_starts_with_digit(procedure_name_with_underscores)
    # the check_if_ID_starts_with_digit(current_ID_cell_value) is just a safety feature for future procedure names/IDs, if it is decided to let them also start with a digit
    return 'SOME_PROCEDURE', 'not inside', 'SOME_PROCEDURE', 'not inside', 'not inside', check_if_ID_starts_with_digit(procedure_name_with_underscores)


def check_if_ID_starts_with_digit(current_ID_cell_value_):
    '''
    checks if id starts with a number and prepends ~ in front of it, if it starts with a number
    :param current_ID_cell_value_:
    :return current_ID_cell_value_:
    '''
    if str(current_ID_cell_value_)[0].isdigit():
        current_ID_cell_value_ = "~" + str(current_ID_cell_value_)
    return current_ID_cell_value_


def check_ENG_string_or_number(current_ENG_cell_value):
    '''
    checks if ENG value is a string or a number
    :param current_ID_cell_value:
    :return current_ID_cell_value:
    '''
    if not str(current_ENG_cell_value).isdigit() and not current_ENG_cell_value == None and not "$" in str(current_ENG_cell_value):
        current_ENG_cell_value = '\"' + str(current_ENG_cell_value) + '\"'
    return current_ENG_cell_value


def alter_values_dependend_on_TYPE(current_TYPE_cell_value, current_RAW_cell_value, current_ENG_cell_value):
    '''
    alters values depending on their type
    :param current_TYPE_cell_value, current_RAW_cell_value, current_ENG_cell_value:
    :return current_RAW_cell_value:
    '''
    if current_TYPE_cell_value == 'Boolean' and not '@$' in str(current_RAW_cell_value) and not '@$' in str(current_ENG_cell_value) and not '$' in str(current_RAW_cell_value):
        current_RAW_cell_value = str(bool(current_RAW_cell_value)).upper()
    return current_RAW_cell_value


def convert_TYPE_from_SCOS_to_MATIS(current_TYPE_cell_value):
    '''
    Converts the SCOS types to MATIS types
    :param current_TYPE_cell_value:
    :return converted_TYPE_value:
    '''
    if str(current_TYPE_cell_value).startswith('Enum') or str(current_TYPE_cell_value).startswith('U8') or str(current_TYPE_cell_value).startswith('U16') or str(current_TYPE_cell_value).startswith('U32') or str(current_TYPE_cell_value).startswith('U64'):
        return 'Unsigned integer'
    elif str(current_TYPE_cell_value).startswith('S8') or str(current_TYPE_cell_value).startswith('S16') or str(current_TYPE_cell_value).startswith('S32') or str(current_TYPE_cell_value).startswith('S64'):
        return 'Signed integer'
    elif str(current_TYPE_cell_value).startswith('Boolean'):
        return 'Boolean'
    elif str(current_TYPE_cell_value).startswith('Float'):
        return 'Real'
    elif str(current_TYPE_cell_value).startswith('Octet Str') or str(current_TYPE_cell_value).startswith('Char Str'):
        return 'String'
    elif str(current_TYPE_cell_value).startswith('Abs Time'):
        return 'Absolute time'
    elif str(current_TYPE_cell_value).startswith('Del Time'):
        return 'Relative time'
    return 'None'


def write_value_plus_add_type_and_description_as_comment(f, identifier_matrix, matrix_iteration, matrix_indicator_operator, current_TYPE_cell_value, current_DESCRIPTION_cell_value, current_RAW_cell_value, current_ENG_cell_value, current_ID_cell_value, indents):
    '''
    Writes the value plus adds the type and description as a comment inside of the PLUTO code.
    :param f, identifier_matrix, matrix_iteration, matrix_indicator_operator, current_TYPE_cell_value, current_DESCRIPTION_cell_value, current_RAW_cell_value, current_ENG_cell_value, current_ID_cell_value, indents:
    :return indents:
    '''
    if current_ENG_cell_value != None:
        write_into_f(f, indents, '{ID_cell}'.format(ID_cell=current_ID_cell_value))
        f.write(' := {ENG}'.format(ENG=str(current_ENG_cell_value).replace('$', '')))
        if identifier_matrix[matrix_iteration + 1][2] == matrix_indicator_operator:
            write_into_f(f, 0, ',')
        if current_RAW_cell_value != None:
            f.write('\t\t//RAW: {RAW}'.format(RAW=str(current_RAW_cell_value).replace('$', '')))
            f.write('\t//TYPE: {TYPE_cell}'.format(TYPE_cell=current_TYPE_cell_value))
        else:
            f.write('\t\t//TYPE: {TYPE_cell}'.format(TYPE_cell=current_TYPE_cell_value))
    else:
#TODO: Remove workaround for "boolean parameter not working in matis (in with arguements raw value of...)"
#Start workaround
        if str(current_TYPE_cell_value).replace(' ', '').startswith('Boolean'):
            if current_RAW_cell_value == 'TRUE':
                current_RAW_cell_value_bool = '\"1\"'
            elif current_RAW_cell_value == 'FALSE':
                current_RAW_cell_value_bool = '\"0\"'
            elif str(current_RAW_cell_value).replace(' ', '').startswith('$'):
                current_RAW_cell_value_bool = str(current_RAW_cell_value).replace('$', '')
            else:
                current_RAW_cell_value_bool = 'NotTRUEorFALSE'
            write_into_f(f, indents,
                         '{ID_cell}'.format(ID_cell=str(current_ID_cell_value).replace('$', '')))
            f.write(' := {RAW_bool}'.format(RAW_bool=str(current_RAW_cell_value_bool).replace('$', '')))
            if identifier_matrix[matrix_iteration + 1][2] == matrix_indicator_operator:
                write_into_f(f, 0, ',')
            f.write('\t\t//TYPE: {TYPE_cell}'.format(TYPE_cell=current_TYPE_cell_value))
            f.write('\t\t//DESCRIPTION: {DESCRIPTION_cell}'.format(DESCRIPTION_cell=current_DESCRIPTION_cell_value))
#End workaround
        else:
            write_into_f(f, indents, 'raw value of {ID_cell}'.format(ID_cell=str(current_ID_cell_value).replace('$', '')))
            f.write(' := {RAW}'.format(RAW=str(current_RAW_cell_value).replace('$', '')))
            if identifier_matrix[matrix_iteration + 1][2] == matrix_indicator_operator:
                write_into_f(f, 0, ',')
            f.write('\t\t//TYPE: {TYPE_cell}'.format(TYPE_cell=current_TYPE_cell_value))
            f.write('\t\t//DESCRIPTION: {DESCRIPTION_cell}'.format(DESCRIPTION_cell=current_DESCRIPTION_cell_value))
    f.write('\n')
    return indents


def check_TM_and_log(f, TC_TM_categories, MIB_TCs_or_TMs, TC_and_TM, SSM, current_RAW_cell_value, current_ID_cell_value, current_DESCRIPTION_cell_value, current_TYPE_cell_value, indents):
    '''
    Writes PLUTO code for TM and writes code to log it.
    :param f, TC_TM_categories, MIB_TCs_or_TMs, TC_and_TM, SSM, current_RAW_cell_value, current_ID_cell_value, current_DESCRIPTION_cell_value, current_TYPE_cell_value, indents:
    :return indents:
    '''
    write_into_f(f, indents, '// DESCRIPTION: {DESCRPITION}, ID: {ID}\n'.format(ID=current_ID_cell_value, DESCRPITION = current_DESCRIPTION_cell_value))
    write_into_f(f, indents, 'if raw_value of ' + str(current_ID_cell_value) + ' of ' + TC_TM_categories + ' of ' + MIB_TCs_or_TMs + ' of ' + TC_and_TM + ' of ' + SSM + ' != ' + str(current_RAW_cell_value).replace('$', '') + ' then\n')
    indents = indent_add(indents)
    write_into_f(f, indents, 'warn \"LOG: FAILURE; ID: {ID}, TYPE: {TYPE}, expected: {RAW}, got: \" + raw_value of {ID} of {TC_TM_categories} of {MIB_TCs_or_TMs} of {TC_and_TM} of {SSM} + \", DESCRIPTION: {DESCRIPTION}\";\n'.format(ID=current_ID_cell_value, TYPE=current_TYPE_cell_value, RAW=current_RAW_cell_value, DESCRIPTION=current_DESCRIPTION_cell_value, TC_TM_categories=TC_TM_categories, MIB_TCs_or_TMs=MIB_TCs_or_TMs, TC_and_TM=TC_and_TM, SSM=SSM))
    indents = indent_remove(indents)
    write_into_f(f, indents, 'end if;\n')
    write_into_f(f, indents, '\n')
    return indents


def check_TM_and_write_into_variable(f, TC_TM_categories, MIB_TCs_or_TMs, TC_and_TM, SSM, current_RAW_cell_value, current_ID_cell_value, current_DESCRIPTION_cell_value, current_TYPE_cell_value, indents):
    '''
    Writes PLUTO code for CHECK TM and writes variable into a file plus logs the value.
    :param f, TC_TM_categories, MIB_TCs_or_TMs, TC_and_TM, SSM, current_RAW_cell_value, current_ID_cell_value, current_DESCRIPTION_cell_value, current_TYPE_cell_value, indents:
    :return indents:
    '''
    write_into_f(f, indents, '// DESCRIPTION: {DESCRPITION}, ID: {ID}\n'.format(ID=current_ID_cell_value, DESCRPITION = current_DESCRIPTION_cell_value))
    write_into_f(f, indents,  str(current_RAW_cell_value).replace(' ', '').replace('@$', '') + ' := ' + str(current_ID_cell_value) + ' of ' + TC_TM_categories + ' of ' + MIB_TCs_or_TMs + ' of ' + TC_and_TM + ' of ' + SSM + ';\n')
    write_into_f(f, indents, '\tlog \"LOG: CHECK TM: VALUE = \" + {ID} of {TC_TM_CATEGORIES} of {MIB_TCS_OR_TMS} of {TC_AND_TM} of {SSM} + \"; DESCRIPTION: {DESCRIPTION}; ID: {ID}\";\n'.format(ID=str(current_ID_cell_value), TC_TM_CATEGORIES=str(TC_TM_categories), MIB_TCS_OR_TMS=str(MIB_TCs_or_TMs), TC_AND_TM=str(TC_and_TM), SSM=str(SSM), DESCRIPTION=str(current_DESCRIPTION_cell_value)))
    return indents


def write_CHECKTM(f, array_declared_variables, TC_TM_categories, MIB_TCs_or_TMs, TC_and_TM, SSM, current_RAW_cell_value, current_ENG_cell_value, current_ID_cell_value, current_DESCRIPTION_cell_value, current_TYPE_cell_value, indents):
    '''
    Writes PLUTO code for CHECK TM command.
    :param f, array_declared_variables, TC_TM_categories, MIB_TCs_or_TMs, TC_and_TM, SSM, current_RAW_cell_value, current_ENG_cell_value, current_ID_cell_value, current_DESCRIPTION_cell_value, current_TYPE_cell_value, indents:
    :return indents:
    '''
    if str(current_ENG_cell_value) != 'None':
        value_cell = str(current_ENG_cell_value)
        flag_ENG_value = 1
        text_raw_value = ''
    else:
        value_cell = str(current_RAW_cell_value)
        flag_ENG_value = 0
        text_raw_value = 'raw_value of '
#Start of cases
    #Value allocation
    if '@' in str(value_cell):
        value_cell = value_cell.replace('@', '').replace('$', '')
        if value_cell not in array_declared_variables:
            value_cell = convert_variable_to_declared_variable_name_with_type(current_TYPE_cell_value, value_cell)
        write_into_f(f, indents, '// DESCRIPTION: {DESCRPITION}, ID: {ID}\n'.format(ID=current_ID_cell_value, DESCRPITION=current_DESCRIPTION_cell_value))
        write_into_f(f, indents, value_cell + ' := ' + text_raw_value + str(current_ID_cell_value) + ' of ' + TC_TM_categories + ' of ' + MIB_TCs_or_TMs + ' of ' + TC_and_TM + ' of ' + SSM + ';\n')
        write_into_f(f, indents, '\tlog \"LOG: CHECK TM ASSIGNMENT: VALUE = \" + {VALUE_CELL} + \"; DESCRIPTION: {DESCRIPTION}; ID: {ID}\";\n\n'.format(
                         ID=str(current_ID_cell_value), VALUE_CELL=value_cell,
                         DESCRIPTION=str(current_DESCRIPTION_cell_value)))
    #Range check
    elif '[' in str(value_cell):
        arg_min = value_cell.split('[', 1)[1].rsplit(',', 1)[0].replace('$', '')
        arg_max = value_cell.split(']', 1)[0].rsplit(',', 1)[1].replace('$', '')
        write_into_f(f, indents, '// DESCRIPTION: {DESCRPITION}, ID: {ID}\n'.format(ID=current_ID_cell_value, DESCRPITION=current_DESCRIPTION_cell_value))
        write_into_f(f, indents, 'if raw_value of ' + str(current_ID_cell_value) + ' of ' + TC_TM_categories + ' of ' + MIB_TCs_or_TMs + ' of ' + TC_and_TM + ' of ' + SSM + ' < ' + arg_min + ' or ' + 'raw_value of ' + str(current_ID_cell_value) + ' of ' + TC_TM_categories + ' of ' + MIB_TCs_or_TMs + ' of ' + TC_and_TM + ' of ' + SSM + ' > ' + arg_max + ' then\n')
        indents = indent_add(indents)
        write_into_f(f, indents,
                     'warn \"LOG: FAILURE; ID: {ID}, TYPE: {TYPE}, expected: {RAW}, got: \" + raw_value of {ID} of {TC_TM_categories} of {MIB_TCs_or_TMs} of {TC_and_TM} of {SSM} + \", DESCRIPTION: {DESCRIPTION}\";\n'.format(
                         ID=current_ID_cell_value, TYPE=current_TYPE_cell_value, RAW=current_RAW_cell_value,
                         DESCRIPTION=current_DESCRIPTION_cell_value, TC_TM_categories=TC_TM_categories,
                         MIB_TCs_or_TMs=MIB_TCs_or_TMs, TC_and_TM=TC_and_TM, SSM=SSM))
        indents = indent_remove(indents)
        write_into_f(f, indents, 'end if;\n')
        write_into_f(f, indents, '\n')
    #Enumeration check
    elif '{' in str(value_cell):
        arg_array = value_cell.split('{', 1)[1].rsplit('}', 1)[0].split(',')
        write_into_f(f, indents, '// DESCRIPTION: {DESCRPITION}, ID: {ID}\n'.format(ID=current_ID_cell_value, DESCRPITION=current_DESCRIPTION_cell_value))
        write_into_f(f, indents, 'if ')
        flag_arg_for_loop = 1
        for arg in arg_array:
            if flag_arg_for_loop == 1:
                write_into_f(f, 0, 'raw_value of ' + str(
                    current_ID_cell_value) + ' of ' + TC_TM_categories + ' of ' + MIB_TCs_or_TMs + ' of ' + TC_and_TM + ' of ' + SSM + ' != ' + arg.replace(' ', '').replace('$', ''))
                flag_arg_for_loop = 0
            else:
                write_into_f(f, 0, ' and raw_value of ' + str(current_ID_cell_value) + ' of ' + TC_TM_categories + ' of ' + MIB_TCs_or_TMs + ' of ' + TC_and_TM + ' of ' + SSM + ' != ' + arg.replace(' ', '').replace('$', ''))
        write_into_f(f, 0, ' then\n')
        indents = indent_add(indents)
        write_into_f(f, indents,
                     'warn \"LOG: FAILURE; ID: {ID}, TYPE: {TYPE}, expected: {RAW}, got: \" + raw_value of {ID} of {TC_TM_categories} of {MIB_TCs_or_TMs} of {TC_and_TM} of {SSM} + \", DESCRIPTION: {DESCRIPTION}\";\n'.format(
                         ID=current_ID_cell_value, TYPE=current_TYPE_cell_value, RAW=current_RAW_cell_value,
                         DESCRIPTION=current_DESCRIPTION_cell_value, TC_TM_categories=TC_TM_categories,
                         MIB_TCs_or_TMs=MIB_TCs_or_TMs, TC_and_TM=TC_and_TM, SSM=SSM))
        indents = indent_remove(indents)
        write_into_f(f, indents, 'end if;\n\n')
    #Size comparison
    elif any(x in str(value_cell) for x in ['>', '>=', '<', '<=']):
        write_into_f(f, indents, '// DESCRIPTION: {DESCRPITION}, ID: {ID}\n'.format(ID=current_ID_cell_value, DESCRPITION=current_DESCRIPTION_cell_value))
        write_into_f(f, indents, 'if raw_value of ' + str(current_ID_cell_value) + ' of ' + TC_TM_categories + ' of ' + MIB_TCs_or_TMs + ' of ' + TC_and_TM + ' of ' + SSM + ' ' + value_cell.replace('$', '') + ' then\n')
        if '$' in str(value_cell):
            write_into_f(f, indents,
                         '\twarn \"LOG: FAILURE; ID: {ID}, TYPE: {TYPE}, expected:\" + {RAW} + \", got: \" + raw_value of {ID} of {TC_TM_categories} of {MIB_TCs_or_TMs} of {TC_and_TM} of {SSM} + \", DESCRIPTION: {DESCRIPTION}\";\n'.format(
                             ID=current_ID_cell_value, TYPE=current_TYPE_cell_value,
                             RAW=str(current_RAW_cell_value).replace('$', ''),
                             DESCRIPTION=current_DESCRIPTION_cell_value, TC_TM_categories=TC_TM_categories,
                             MIB_TCs_or_TMs=MIB_TCs_or_TMs, TC_and_TM=TC_and_TM, SSM=SSM))
        else:
            write_into_f(f, indents,
                         '\twarn \"LOG: FAILURE; ID: {ID}, TYPE: {TYPE}, expected: {RAW}, got: \" + raw_value of {ID} of {TC_TM_categories} of {MIB_TCs_or_TMs} of {TC_and_TM} of {SSM} + \", DESCRIPTION: {DESCRIPTION}\";\n'.format(
                             ID=current_ID_cell_value, TYPE=current_TYPE_cell_value, RAW=str(current_RAW_cell_value),
                             DESCRIPTION=current_DESCRIPTION_cell_value, TC_TM_categories=TC_TM_categories,
                             MIB_TCs_or_TMs=MIB_TCs_or_TMs, TC_and_TM=TC_and_TM, SSM=SSM))
        write_into_f(f, indents, 'end if;\n\n')
    #TM check based on variable value
    elif str(value_cell).replace(' ', '').startswith('$'):
        write_into_f(f, indents, '// DESCRIPTION: {DESCRPITION}, ID: {ID}\n'.format(ID=current_ID_cell_value, DESCRPITION=current_DESCRIPTION_cell_value))
        write_into_f(f, indents, 'if raw_value of ' + str(current_ID_cell_value) + ' of ' + TC_TM_categories + ' of ' + MIB_TCs_or_TMs + ' of ' + TC_and_TM + ' of ' + SSM + ' = ' + value_cell.replace('$', '') + ' then\n')
        write_into_f(f, indents,
                     '\twarn \"LOG: FAILURE; ID: {ID}, TYPE: {TYPE}, expected:\" + {RAW} + \", got: \" + raw_value of {ID} of {TC_TM_categories} of {MIB_TCs_or_TMs} of {TC_and_TM} of {SSM} + \", DESCRIPTION: {DESCRIPTION}\";\n'.format(
                         ID=current_ID_cell_value, TYPE=current_TYPE_cell_value, RAW=str(current_RAW_cell_value).replace('$', ''),
                         DESCRIPTION=current_DESCRIPTION_cell_value, TC_TM_categories=TC_TM_categories,
                         MIB_TCs_or_TMs=MIB_TCs_or_TMs, TC_and_TM=TC_and_TM, SSM=SSM))
        write_into_f(f, indents, 'end if;\n\n')
    #TM check based on fixed value
    else:
        write_into_f(f, indents, '// DESCRIPTION: {DESCRPITION}, ID: {ID}\n'.format(ID=current_ID_cell_value, DESCRPITION=current_DESCRIPTION_cell_value))
        write_into_f(f, indents, 'if raw_value of ' + str(current_ID_cell_value) + ' of ' + TC_TM_categories + ' of ' + MIB_TCs_or_TMs + ' of ' + TC_and_TM + ' of ' + SSM + ' = ' + value_cell.replace('$', '') + ' then\n')
        write_into_f(f, indents,
                     '\twarn \"LOG: FAILURE; ID: {ID}, TYPE: {TYPE}, expected: {RAW}, got: \" + raw_value of {ID} of {TC_TM_categories} of {MIB_TCs_or_TMs} of {TC_and_TM} of {SSM} + \", DESCRIPTION: {DESCRIPTION}\";\n'.format(
                         ID=current_ID_cell_value, TYPE=current_TYPE_cell_value, RAW=str(current_RAW_cell_value).replace('$', ''),
                         DESCRIPTION=current_DESCRIPTION_cell_value, TC_TM_categories=TC_TM_categories,
                         MIB_TCs_or_TMs=MIB_TCs_or_TMs, TC_and_TM=TC_and_TM, SSM=SSM))
        write_into_f(f, indents, 'end if;\n\n')
    return indents


def convert_variable_to_declared_variable_name_with_type(current_TYPE_cell_value, value_cell):
    '''
    Converts the variable VAL to its specific type with its name.
    :param current_TYPE_cell_value, value_cell:
    :return value_cell:
    '''
    if current_TYPE_cell_value == 'Unsigned integer':
        type_addon = 'UI'
    elif current_TYPE_cell_value == 'Signed integer':
        type_addon = 'SI'
    elif current_TYPE_cell_value == 'Boolean':
        type_addon = 'BOOL'
    elif current_TYPE_cell_value == 'Real':
        type_addon = 'REAL'
    elif current_TYPE_cell_value == 'String':
        type_addon = 'STR'
    elif current_TYPE_cell_value == 'Absolute time':
        type_addon = 'ABST'
    elif current_TYPE_cell_value == 'Relative time':
        type_addon = 'RELT'
    else:
        type_addon = ''
    value_cell = 'VAL' + '_' + type_addon
    return value_cell


def write_DECLARE_TM_CHECK_VARIABLES(f, array_declared_variables, flag_array_TM_CHECK_VARIABLES, identifier_matrix, matrix_indicator_operator_old, matrix_indicator_operator, current_ID_cell_value, current_DESCRIPTION_cell_value, current_TYPE_cell_value, matrix_iteration, indents, flag_some_variable_declared):
    '''
    Writes the variables into the DECLARE method in PLUTO where all variables
    are declared and adds the correct name of VAL_... .
    :param f, array_declared_variables, flag_array_TM_CHECK_VARIABLES, identifier_matrix, matrix_indicator_operator_old, matrix_indicator_operator, current_ID_cell_value, current_DESCRIPTION_cell_value, current_TYPE_cell_value, matrix_iteration, indents, flag_some_variable_declared:
    :return indents, flag_array_TM_CHECK_VARIABLES, array_declared_variables, flag_some_variable_declared:
    '''
    flag_array_TM_CHECK_VARIABLES_BUFFER = flag_array_TM_CHECK_VARIABLES.copy()
    if str(current_TYPE_cell_value) == 'Unsigned integer' and flag_array_TM_CHECK_VARIABLES[0] == 1:
        VAL_NAME = 'VAL_UI'
        flag_array_TM_CHECK_VARIABLES[0] = 0
    elif str(current_TYPE_cell_value) == 'Signed integer' and flag_array_TM_CHECK_VARIABLES[1] == 1:
        VAL_NAME = 'VAL_SI'
        flag_array_TM_CHECK_VARIABLES[1] = 0
    elif str(current_TYPE_cell_value) == 'Boolean' and flag_array_TM_CHECK_VARIABLES[2] == 1:
        VAL_NAME = 'VAL_BOOL'
        flag_array_TM_CHECK_VARIABLES[2] = 0
    elif str(current_TYPE_cell_value) == 'Real' and flag_array_TM_CHECK_VARIABLES[3] == 1:
        VAL_NAME = 'VAL_REAL'
        flag_array_TM_CHECK_VARIABLES[3] = 0
    elif str(current_TYPE_cell_value) == 'String' and flag_array_TM_CHECK_VARIABLES[4] == 1:
        VAL_NAME = 'VAL_STR'
        flag_array_TM_CHECK_VARIABLES[4] = 0
    elif str(current_TYPE_cell_value) == 'Absolute time' and flag_array_TM_CHECK_VARIABLES[5] == 1:
        VAL_NAME = 'VAL_ABST'
        flag_array_TM_CHECK_VARIABLES[5] = 0
    elif str(current_TYPE_cell_value) == 'Relative time' and flag_array_TM_CHECK_VARIABLES[6] == 1:
        VAL_NAME = 'VAL_RELT'
        flag_array_TM_CHECK_VARIABLES[6] = 0
    else:
        VAL_NAME = 'None'
    if flag_array_TM_CHECK_VARIABLES != flag_array_TM_CHECK_VARIABLES_BUFFER:
        if array_declared_variables != [] and flag_array_TM_CHECK_VARIABLES != [1, 1, 1, 1, 1, 1, 1] and flag_some_variable_declared != 1:
            write_into_f(f, 0, ',\n')
        else:
            write_into_f(f, 0, '\n')
            flag_some_variable_declared = 0
        write_into_f(f, indents, 'variable {VAL} of type {TYPE}'.format(VAL=VAL_NAME, TYPE=current_TYPE_cell_value))
        array_declared_variables.append(VAL_NAME)
    return indents, flag_array_TM_CHECK_VARIABLES, array_declared_variables, flag_some_variable_declared


def create_identifier_matrix(ws_procedure):
    '''
    Creates identifier matrix which is used all over this code
    :param ws_procedure:
    :return identifier_matrix:
    '''
    ## START MATRIX PREPARATIONS
    # get all rows_procedure with the start of a new operation
    new_operation_row_numbers = get_operations_captions_row_number(ws_procedure)
    # iterate through the different operation topics
    iterationNumber = 0
    identifier_matrix = []
    identifier_matrix.append(['row_number', 'indicator1', 'indicator_operator'])
    for row in new_operation_row_numbers[:-1]:
        # add row to identify operation-topic switch
        current_OPERATIONS_cell_value = ws_procedure[
            _OPERATIONS_COLUMN + '{rows_procedure}'.format(rows_procedure=row)].value
        identifier_matrix.append([row, 'NEW_OPERATION_STEP', str(current_OPERATIONS_cell_value)])
        # print all operation topics
        print(ws_procedure[_OPERATIONS_COLUMN + '{rows_procedure}'.format(rows_procedure=row)].value)
        identifier_matrix = iterating_over_operation_topic(ws_procedure, new_operation_row_numbers, iterationNumber,
                                                           identifier_matrix)
        iterationNumber += 1
    # add row to identify last operation-topic switch (end of procedure)
    identifier_matrix.append([new_operation_row_numbers[len(new_operation_row_numbers) - 1], 'NEW_OPERATION_STEP', str(
        ws_procedure[
            _OPERATIONS_COLUMN + '{rows_procedure}'.format(rows_procedure=new_operation_row_numbers[-1])].value)])
    print(ws_procedure[_OPERATIONS_COLUMN + '{rows_procedure}'.format(
        rows_procedure=new_operation_row_numbers[len(new_operation_row_numbers) - 1])].value)
    # identifier_matrix = generate_indicator_operator_identifier_matrix(ws_procedure, identifier_matrix)
    #print(identifier_matrix)
    return identifier_matrix


def get_current_row_cells(ws_procedure, matrix_row_number):
    '''
    Extracts the current row cell and outputs its values
    :param ws_procedure, matrix_row_number:
    :return current_STEP_cell_value, current_OPERATIONS_cell_value, current_ID_cell_value, current_DESCRIPTION_cell_value, current_TYPE_cell_value, current_RAW_cell_value, current_ENG_cell_value, current_UNIT_cell_value:
    '''
    current_STEP_cell_value = ws_procedure[_STEP_COLUMN + '{rows}'.format(rows=matrix_row_number)].value
    current_OPERATIONS_cell_value = ws_procedure[_OPERATIONS_COLUMN + '{rows}'.format(rows=matrix_row_number)].value
    current_ID_cell_value = ws_procedure[_ID_COLUMN + '{rows}'.format(rows=matrix_row_number)].value
    current_ID_cell_value = check_if_ID_starts_with_digit(current_ID_cell_value)
    current_DESCRIPTION_cell_value = ws_procedure[_DESCRIPTION_COLUMN + '{rows}'.format(rows=matrix_row_number)].value
    current_TYPE_cell_value = ws_procedure[_TYPE_COLUMN + '{rows}'.format(rows=matrix_row_number)].value
    current_TYPE_cell_value = convert_TYPE_from_SCOS_to_MATIS(current_TYPE_cell_value)
    current_ENG_cell_value = ws_procedure[_ENG_COLUMN + '{rows}'.format(rows=matrix_row_number)].value
    current_ENG_cell_value = check_ENG_string_or_number(current_ENG_cell_value)
    current_RAW_cell_value = ws_procedure[_RAW_COLUMN + '{rows}'.format(rows=matrix_row_number)].value
    current_RAW_cell_value = alter_values_dependend_on_TYPE(current_TYPE_cell_value, current_RAW_cell_value, current_ENG_cell_value)
    current_UNIT_cell_value = ws_procedure[_UNIT_COLUMN + '{rows}'.format(rows=matrix_row_number)].value
    return current_STEP_cell_value, current_OPERATIONS_cell_value, current_ID_cell_value, current_DESCRIPTION_cell_value, current_TYPE_cell_value, current_RAW_cell_value, current_ENG_cell_value, current_UNIT_cell_value


def indent_add(indents):
    '''
    Adds one indent
    :param indents:
    :return indents:
    '''
    indents += 1
    return indents


def indent_remove(indents):
    '''
    Removes one indent
    :param indents:
    :return indents:
    '''
    if indents > 0:
        indents -= 1
    return indents


def write_into_f(f, number_of_indents, text):
    '''
    Writes input into file with correct indents.
    :param f, number_of_indents, text:
    :return :
    '''
    indents = ''
    for addingIndent in range(0, number_of_indents):
        indents = indents + '\t'
    f.write('{CONTENT}'.format(CONTENT=str(indents+text)))
    return


## #############################################
# START OF PROGRAMME
## #############################################

# GLOBAL VARIABLE DEFINITIONS
# file names
global _EXCEL_FILE_NAME # e.g. 'R-ADC-N210_Activate_cADCS_IDLE_mode_v4.xlsx'

global _FILE_NAME #str(_EXCEL_FILE_NAME)[0:10].replace('-', '_')   #'R_ADC_N210_v4.txt'
# columns and cells for front page work sheet
_PROCEDURE_TITLE_CELL = 'D3'
_PROCEDURE_ID_CELL = 'D4'
_START_COLUMN_FRONT_PAGE = 'A'
_FILLER_COLUMN1_FRONT_PAGE = 'B'
_FILLER_COLUMN2_FRONT_PAGE = 'C'
_INFORMATION_COLUMN_FRONT_PAGE = 'D'
_END_COLUMN_FRONT_PAGE = 'E'
# columns for procedure worksheet
_STEP_COLUMN = 'A'
_OPERATIONS_COLUMN = 'B'
_ID_COLUMN = 'C'
_DESCRIPTION_COLUMN = 'D'
_TYPE_COLUMN = 'E'
_RAW_COLUMN = 'F'
_ENG_COLUMN = 'G'
_UNIT_COLUMN = 'H'
_DISPLAY_COLUMN = 'I'
_CONFIRMATION_COLUMN = 'Q'
_COLOR_DIVIDING_OPERATION_STEPS = 'FF92CDDC'
# global variable for unknown commands and comments
_BUFFER_FOR_TEXT_FOR_UNKNOWN_COMMANDS = ''
# known operations which can be directly implemented into the code
#TODO: ask Daniela for more known parameters
_KNOWN_OPERATIONS_PARAMETER = ['SEND', 'SEND_', 'SENDTIMETAG', 'SENDTIMETAG_', 'SENDANDCHECKTCV', 'SENDANDCHECKTCV_', 'CHECKTM', 'CHECKTM_', 'DECLAREVARIABLES', 'SELECTCASE', 'CASE:', '$', 'CASEELSE', 'ENDCASE', 'IF', 'ELSEIF', 'ELSE', 'THEN', 'ENDIF', 'CALLPROCEDURE', 'THENRETURN', 'EXECUTEINTERMINALONMCSMACHINE', 'CALLENGINEER', 'WAIT']

#TODO: Make it pretty
# this function has been placed here, since it belongs here (flow vise)
def main_function(input_file, output_file):
    '''
    Has all functionality of writing the PLUTO code and also checks the file for forbidden characters.
    :param input_file, output_file:
    :return :
    '''
    ## START: PYTHON
    # load excel f
    wb = op.load_workbook(str(_EXCEL_FILE_NAME), data_only=True)
    #  open certain sheet (tab) in excel f
    ws_front_page = wb['Front Page']
    ws_procedure = wb['Procedure']
    write_DATE_of_autogeneration_and_initials()
    write_front_page_documentation_as_comment_into_f(ws_front_page) # also closes and opens the file ->therefore new f (file-pointer)
    print('operations: ', get_operations_captions_row_number(ws_procedure))
    identifier_matrix = create_identifier_matrix(ws_procedure)
    print(identifier_matrix)
    ## CREATE DICTIONARY FOR PARAMETER ASSIGNMENT
    tree_structure_params_repository = create_parameter_dictionary()
    ## GENERATE PLUTO CODE
    generate_code(ws_front_page, ws_procedure, _BUFFER_FOR_TEXT_FOR_UNKNOWN_COMMANDS, identifier_matrix, tree_structure_params_repository, asdf_list)
    ## DELETE FORBIDDEN CHARACTERS
    check_file_for_forbidden_characters(_FILE_NAME)
    check_file_for_empty_steps_and_delete(_FILE_NAME)
    #print(identifier_matrix)
    return

asdf_list = []
rootdir = os.getcwd()
list_of_excelsheet_paths =[]
for subdir, dirs, files in os.walk(rootdir + '\\Excel\\'):
    for file in files:
        #print os.path.join(subdir, file)
        filepath = subdir + os.sep + file
        if filepath.endswith(".xlsx") and '\\old\\' not in filepath and '~' not in filepath:
            list_of_excelsheet_paths.append(str(filepath).split('\\Excel\\', 1)[1])
print(list_of_excelsheet_paths)
iteration = 0
shutil.rmtree('generated_MATIS_Files')
os.makedirs('generated_MATIS_Files')
for file in list_of_excelsheet_paths:
    iteration += 1
    input_file_path = 'Excel\\' + file
    directory = 'generated_MATIS_Files\\' + file.rsplit('\\', 1)[0]
    if not os.path.exists(directory):
        os.makedirs(directory)
    #myfunction('Excel\\' + str(file), output_location)
    output_file_ID_name = file.rsplit('\\', 1)[1].split('-', 1)
    output_file_ID_name[0] = output_file_ID_name[0].replace('\\', '').replace('-', '_') + '_'
    print('outputfilename[0] = ', output_file_ID_name[0])
    output_file_ID_name[1] = output_file_ID_name[1].replace('\\', '')[0:8].replace('-', '_')
    print('outputfilename[1] = ', output_file_ID_name[1])
    output_file_ID_name = output_file_ID_name[0] + output_file_ID_name[1]
    print('outputfilename = ', output_file_ID_name)
    output_file_path = str(directory) + '\\' + output_file_ID_name + '.pluto'
    print('input: ' + input_file_path)
    print('output: ', output_file_path)
    _EXCEL_FILE_NAME = input_file_path
    _FILE_NAME = output_file_path
    main_function(input_file_path, output_file_path)

## Start: Print files which contain "@$", can be changed in code above (search for asdf_list) to find and print other
# files with certain pattern/characteristics. Especially useful when a certain pattern has to be changed in excel procedure.
asdf_old = ''
for asdf in asdf_list:
    if asdf != asdf_old:
        print(asdf)
    asdf_old = asdf
    ## #############################################
    # END OF PROGRAMME
    ## #############################################