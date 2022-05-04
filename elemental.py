import os
import cantera as ct
import numpy as np
import pandas as pd

import xlwt
from xlwt import Workbook
import openpyxl as xl

import xlsxwriter

from rdkit import Chem
from rdkit.Chem.rdMolDescriptors import CalcMolFormula

# Create Excel workbook
wb = xlsxwriter.Workbook("C:\\Users\\opbir\\Documents\\ML_trainingData_element.xlsx");

# Read all training data files
directory = 'C:\\Program Files\\Cantera\\data\\ML datasets';
ct.add_directory('C:\\Program Files\\Cantera\\data\\ML datasets');

elements = ['C','H','O','N'];

sheet = wb.add_worksheet();
sheet.write(0,0,'Element');
sheet.write(1,0,'C');
sheet.write(2,0,'H');
sheet.write(3,0,'O');
sheet.write(4,0,'N');
sheet.write(5,0,'AR');

uniqueSpecies = [];
uniqueSpecies_sym = [];

count1 = 1;
for filename in os.listdir(directory):
    f = os.path.join(directory,filename)         # full file directory
    if os.path.isfile(f):
        gas = ct.Solution(filename);

        # Get all unique species from all datasets
        for x in gas.species():
            species = x.name;
            if species not in uniqueSpecies:
                uniqueSpecies.append(species);
                sheet.write(0,count1,species);
                count1 = count1 + 1;
                for i in elements:
                    sheet_loc = elements.index(i)+1;
                    n = gas.n_atoms(species,i);
                    sheet.write(sheet_loc,count1-1,n);
                if species == 'AR':
                    sheet.write(5,count1-1,1);
print(uniqueSpecies);
print(len(uniqueSpecies))

fileExcel = pd.ExcelFile("C:\\Users\\opbir\\Documents\\ML_trainingData_element.xlsx");   
for sheetName in fileExcel.sheet_names:
    df = pd.read_excel("C:\\Users\\opbir\\Documents\\ML_trainingData_element.xlsx");
    df.fillna(0,inplace=True);

    df.to_excel('C:\\Users\\opbir\\Documents\\ML_trainingData_element_df.xlsx');
    print(df);

wb.close();