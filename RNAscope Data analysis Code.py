#!/usr/bin/env python3
# RNAscope data extraction code
"""

RNAscope data extraction Code
Brenhouse Lab
By: Abigail Parakoyi
Commission by: Lauren Granata
09/07/22

"""
import pandas as pd
import os


def main():
    """ Business Logic"""
    # opens the Excel file and allows read access to it
    root_dir = 'MPOA Data 20.01-20.10'
    print('running')
    # print(list(os.listdir(root_dir)))
    for file in os.listdir(root_dir):
        os.path.join(directory, file)
        data = get_filehandle(root_dir+'/'+file)
        # path = "/Users/abigailparakoyi/Desktop/LabWork/20.01.xlsx"  # hardcode the absolute path to data sheet
        # data = get_filehandle(path)
            # print(data) # keep for testing purposes
        data = data.fillna(0)   # open the data and replace all blanks with zeros
        dictionary = make_dict(data)
        df = pd.DataFrame(data=dictionary, index=['Subject Number'])
        df.to_excel("test3.xlsx")   # Change name to desired excel sheet name
        print("converted to excel")
    # print(dictionary)


def get_filehandle(file):
    """
    Get the df for the inputted file
    :param file: Desired RNAscope file
    :return: df of desired RNAscope file
    """
    try:
        df = pd.read_excel(file, dtype={'Image': str, 'Name': str, 'Class': str,
                                        'Parent': str, 'Cell: GnRH mean': float, 'Cell: GnRH std dev': float,
                                        'Cell: Kiss1 mean': float, 'Cell: Kiss1 std dev': float, 'Cell: Crhr1 mean': float,
                                        'Cell: Crhr1 std dev': float, 'Subcellular: Channel 1: Num spots estimated': float and str,
                                        'Subcellular: Channel 2: Num spots estimated': float and str, 'Subcellular: Channel 4: Num spots estimated': float and str,
                                        'Subcellular cluster: Channel 4: Mean channel intensity': float, 'Subcellular cluster: Channel 2: Mean channel intensity': float,
                                        'Subcellular cluster: Channel 1: Mean channel intensity': float}, engine='openpyxl')   # read through and assign data sheet to variable
        # print(df['Image']) # this is keyed on columns3
        return df
    except OSError as error:
        print(f'{file} cannot be opened')
        raise error
# print(subject_data)     # Check if data reads from txt file
# print(subject_data.head)       # Checking the first 10 rows
# print(subject_data.loc[0])    # Print the 1st row
# print(subject_data.dtypes)      # looking at the data types in each column


def make_dict(df):
    """
    Make dictionary of data creating one file into a row
    :param df: the fil of interest
    :return: one dictionary for each animal
    """
    final_dict = {'Subject Number': '', 'Total Cell Count': 0, 'GnRH Cell Count': 0, 'Kiss1 Cell Count': 0, 'Crhr1 Cell Count': 0,
                  'GnRH: Crhr1 Cell Count': 0, 'Kiss1: Crhr1 Cell Count': 0, 'GnRH Average intensity/cell': 0, 'Kiss1 Average intensity/cell': 0,
                  'Crhr1 Average Intensity/cell': 0, 'GnRH: Crhr1 Average Intensity/cell': 0, 'Kiss1: Crhr1 Average Intensity/cell': 0,
                  'GnRH Average intensity/cluster': 0, 'Kiss1 Average intensity/cluster': 0, 'Crhr1 Average Intensity/cluster': 0,
                  'GnRH: Crhr1 Average Intensity/cluster': 0, 'Kiss1: Crhr1 Average Intensity/cluster': 0}
    for index, row in df.iterrows():
        sbj = row['Image'].split('_')    # Image name in row and split on _
        final_dict['Subject Number'] = sbj[0]
        if row['Class'] == 'GnRH':
            final_dict['GnRH Average intensity/cell'] += row['Cell: GnRH mean']
        if row['Class'] == 'Kiss1':
            final_dict['Kiss1 Average intensity/cell'] += row['Cell: Kiss1 mean']
        if row['Class'] == 'Crhr1':
            final_dict['Crhr1 Average Intensity/cell'] += row['Cell: Crhr1 mean']
        if row['Class'] == 'GnRH: Crhr1':
            final_dict['GnRH: Crhr1 Average Intensity/cell'] += row['Cell: Crhr1 mean'] + row['Cell: GnRH mean']
        if row['Class'] == 'Kiss1: Crhr1':
            final_dict['Kiss1: Crhr1 Average Intensity/cell'] += row['Cell: Crhr1 mean'] + row['Cell: Kiss1 mean']
        if row['Class'] == 'Subcellular spot: Channel 1 object' or 'Subcellular cluster: Channel 1 object' and row['Parent'] == 'GnRH':
            final_dict['GnRH Average intensity/cluster'] += row['Subcellular cluster: Channel 1: Mean channel intensity']
        if row['Class'] == 'Subcellular spot: Channel 2 object' or 'Subcellular cluster: Channel 2 object' and row['Parent'] == 'Kiss1':
            final_dict['Kiss1 Average intensity/cluster'] += row['Subcellular cluster: Channel 2: Mean channel intensity']
        if row['Class'] == 'Subcellular spot: Channel 4 object' or 'Subcellular cluster: Channel 4 object' and row['Parent'] == 'Crhr1':
            final_dict['Crhr1 Average Intensity/cluster'] += row['Subcellular cluster: Channel 4: Mean channel intensity']
        if row['Class'] == 'Subcellular spot: Channel 1 object' or 'Subcellular cluster: Channel 1 object' and 'Subcellular spot: Channel 4 object' or 'Subcellular cluster: Channel 4 object' and row['Parent'] == 'GnRH: Crhr1':
            final_dict['GnRH: Crhr1 Average Intensity/cluster'] += row['Subcellular cluster: Channel 1: Mean channel intensity'] + row['Subcellular cluster: Channel 4: Mean channel intensity']
        if row['Class'] == 'Subcellular spot: Channel 2 object' or 'Subcellular cluster: Channel 2 object' and 'Subcellular spot: Channel 4 object' or 'Subcellular cluster: Channel 4 object' and row['Parent'] == 'Kiss1: Crhr1':
            final_dict['Kiss1: Crhr1 Average Intensity/cluster'] += row['Subcellular cluster: Channel 2: Mean channel intensity'] + row['Subcellular cluster: Channel 4: Mean channel intensity']
        if row['Name'] == 'PathCellObject' and row['Class'] == 'GnRH' or 'Kiss1' or 'Crhr1':
            final_dict['Total Cell Count'] += 1
        if row['Class'] == 'GnRH':
            final_dict['GnRH Cell Count'] += 1
        if row['Class'] == 'Kiss1':
            final_dict['Kiss1 Cell Count'] += 1
        if row['Class'] == 'Crhr1':
            final_dict['Crhr1 Cell Count'] += 1
        if row['Class'] == 'GnRH: Crhr1':
            final_dict['GnRH: Crhr1 Cell Count'] += 1
        if row['Class'] == 'Kiss1: Crhr1':
            final_dict['Kiss1: Crhr1 Cell Count'] += 1
    subcellular_count_GnRH = 0
    subcellular_count_Kiss1 = 0
    subcellular_count_Crhr1 = 0
    subcellular_count_GnRH_Crhr1 = 0
    subcellular_count_Kiss1_Crhr1 = 0
    for index, row in df.iterrows():
        if row['Class'] == 'Subcellular spot: Channel 1 object' or 'Subcellular cluster: Channel 1 object' and row['Parent'] == 'GnRH':
            subcellular_count_GnRH += 1
        if row['Class'] == 'Subcellular spot: Channel 2 object' or 'Subcellular cluster: Channel 2 object' and row['Parent'] == 'Kiss1':
            subcellular_count_Kiss1 += 1
        if row['Class'] == 'Subcellular spot: Channel 4 object' or 'Subcellular cluster: Channel 4 object' and row['Parent'] == 'Crhr1':
            subcellular_count_Crhr1 += 1
        if row['Class'] == 'Subcellular spot: Channel 1 object' or 'Subcellular cluster: Channel 1 object' and 'Subcellular spot: Channel 4 object' or 'Subcellular cluster: Channel 4 object' and row['Parent'] == 'GnRH: Crhr1':
            subcellular_count_GnRH_Crhr1 += 1
        if row['Class'] == 'Subcellular spot: Channel 2 object' or 'Subcellular cluster: Channel 2 object' and 'Subcellular spot: Channel 4 object' or 'Subcellular cluster: Channel 4 object' and row['Parent'] == 'Kiss1: Crhr1':
            subcellular_count_Kiss1_Crhr1 += 1
    final_dict['GnRH Average intensity/cell'] /= final_dict['GnRH Cell Count']
    final_dict['Kiss1 Average intensity/cell'] /= final_dict['Kiss1 Cell Count']
    final_dict['Crhr1 Average Intensity/cell'] /= final_dict['Crhr1 Cell Count']
    final_dict['GnRH: Crhr1 Average Intensity/cell'] /= final_dict['GnRH: Crhr1 Cell Count']
    final_dict['Kiss1: Crhr1 Average Intensity/cell'] /= final_dict['Kiss1: Crhr1 Cell Count']
    final_dict['GnRH Average intensity/cluster'] /= subcellular_count_GnRH
    final_dict['Kiss1 Average intensity/cluster'] /= subcellular_count_Kiss1
    final_dict['Crhr1 Average Intensity/cluster'] /= subcellular_count_Crhr1
    final_dict['GnRH: Crhr1 Average Intensity/cluster'] /= subcellular_count_GnRH_Crhr1
    final_dict['Kiss1: Crhr1 Average Intensity/cluster'] /= subcellular_count_Kiss1_Crhr1
        # print(final_dict) # for testing purposes
    return final_dict


if __name__ == '__main__':
    main()
