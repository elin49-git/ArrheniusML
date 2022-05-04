import os
import cantera as ct
import numpy as np
import pandas as pd

import xlwt
from xlwt import Workbook
import openpyxl as xl
import xlsxwriter

# Create Excel workbook
wb = xlsxwriter.Workbook("C:\\Users\\opbir\\Documents\\ML_trainingData_moleFraction.xlsx");

# Read all training data files
directory = 'C:\\Program Files\\Cantera\\data\\ML datasets';
ct.add_directory('C:\\Program Files\\Cantera\\data\\ML datasets');

# Initializations
dP = float('inf');
uniqueSpecies = [];
MW = [];

for filename in os.listdir(directory):
    f = os.path.join(directory,filename)         # full file directory
    if os.path.isfile(f):
        gas = ct.Solution(filename);

        # Get all unique species from all datasets
        for x in gas.species():
            species = x.name;
            gasMW = gas.molecular_weights;
            if species not in uniqueSpecies:
                loc = 0;
                while gas.species()[loc].name != species:
                    loc = loc + 1;

                uniqueSpecies.append(species);
                MW.append(gasMW[loc]);

print(uniqueSpecies);
print(MW);
   
for filename in os.listdir(directory):
    f = os.path.join(directory,filename)         # full file directory
    if os.path.isfile(f):
        gas = ct.Solution(filename);

        fileNo = 0;
        sheet = wb.add_worksheet(filename);
        sheet.write(0,0,'Reactions');

        count1 = 0;
        for speciesName in uniqueSpecies:
            sheet.write(0,count1+1,speciesName);
            count1 = count1 + 1;

        indx = [i for i, r in enumerate(gas.reactions())];
        for i in indx:
            print(gas.reaction(i).equation);                 # Reaction
            sheet.write(i+1,fileNo,gas.reaction(i).equation);
           
            # For each species in a reaction i
            count = 0;
            for sp in gas.species():
                findSpecies = sp.name;                
                location = uniqueSpecies.index(findSpecies)
                
                reactantCoeff = -1*gas.reactant_stoich_coeff(findSpecies,i);       # Reactant: -
                productCoeff = gas.product_stoich_coeff(findSpecies,i);            # Product: +

                if reactantCoeff != 0:
                    sheet.write(i+1,location+1,reactantCoeff*MW[location]);

                if productCoeff != 0:
                    sheet.write(i+1,location+1,productCoeff*MW[location]);

                count = count + 1;
                
fileExcel = pd.ExcelFile("C:\\Users\\opbir\\Documents\\ML_trainingData_moleFraction.xlsx");   
print(fileExcel.sheet_names)

for sheetName in fileExcel.sheet_names:
    df = pd.read_excel("C:\\Users\\opbir\\Documents\\ML_trainingData_moleFraction.xlsx",sheet_name = sheetName);
    df.fillna(0,inplace=True);

    df.to_excel('C:\\Users\\opbir\\Documents\\ML_trainingData__moleFraction_'+sheetName[:-4]+'.xlsx');
    print(df);

wb.close();
