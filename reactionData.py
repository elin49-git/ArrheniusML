import os
import cantera as ct
import numpy as np
import pandas as pd

import xlwt
from xlwt import Workbook
import openpyxl as xl

import xlsxwriter

# Create Excel workbook
wb = xlsxwriter.Workbook("C:\\Users\\opbir\\Documents\\ML_trainingData.xlsx");

# Read all training data files
directory = 'C:\\Program Files\\Cantera\\data\\ML datasets';
ct.add_directory('C:\\Program Files\\Cantera\\data\\ML datasets');

# Initializations
dP = float('inf');
uniqueSpecies = [];

for filename in os.listdir(directory):
    f = os.path.join(directory,filename)         # full file directory
    if os.path.isfile(f):
        gas = ct.Solution(filename);

        # Get all unique species from all datasets
        for x in gas.species():
            species = x.name;
            if species not in uniqueSpecies:
                uniqueSpecies.append(species);

print(uniqueSpecies);
    
for filename in os.listdir(directory):
    f = os.path.join(directory,filename)         # full file directory
    if os.path.isfile(f):
        gas = ct.Solution(filename);

        fileNo = 0;
        sheet = wb.add_worksheet(filename);

        sheet.write(0,0,'Reactions');
        sheet.write(0,1,'A');
        sheet.write(0,2,'b');
        sheet.write(0,3,'E');
        sheet.write(0,4,'Eq. Const (K)');
        sheet.write(0,5,'f_rateConst');
        sheet.write(0,6,'dEnthalpy');
        sheet.write(0,7,'dEntropy');
        sheet.write(0,8,'dGibbs');

        count1 = 0;
        for speciesName in uniqueSpecies:
            sheet.write(0,9+count1,speciesName);
            count1 = count1 + 1;

        indx = [i for i, r in enumerate(gas.reactions())];
        for i in indx:
            print(gas.reaction(i).equation);                 # Reaction
            sheet.write(i+1,fileNo,gas.reaction(i).equation);

            # For each reaction
            eqConst = gas.equilibrium_constants[i];
            f_rateConst = gas.forward_rate_constants[i];
            dEnthalpy = gas.delta_enthalpy[i];
            dEntropy = gas.delta_entropy[i];
            dGibbs = gas.delta_gibbs[i];

            sheet.write(i+1,fileNo+4,eqConst);
            sheet.write(i+1,fileNo+5,f_rateConst);
            sheet.write(i+1,fileNo+6,dEnthalpy);
            sheet.write(i+1,fileNo+7,dEntropy);
            sheet.write(i+1,fileNo+8,dGibbs);
            
            # For each species in a reaction i
            count = 0;
            for sp in gas.species():
                findSpecies = sp.name;
                location = uniqueSpecies.index(findSpecies);      
                
                reactantCoeff = -1*gas.reactant_stoich_coeff(findSpecies,i);       # Reactant: -
                productCoeff = gas.product_stoich_coeff(findSpecies,i);            # Product: +

                if reactantCoeff != 0:
                    sheet.write(i+1,location+9,reactantCoeff);
                if productCoeff != 0:
                    sheet.write(i+1,location+9,productCoeff);

                count = count + 1;

                print(filename)
                print(i)
                print(sp)
                
            rxn = gas.reaction(i);

            if gas.reaction_type(i) == 1 or gas.reaction_type(i) == 2:    # If single or three-body reaction
                arr = rxn.rate; 
                A = arr.pre_exponential_factor;   # A
                b = arr.temperature_exponent;     # b
                E = arr.activation_energy;        # E

                sheet.write(i+1,fileNo+1,A);
                sheet.write(i+1,fileNo+2,b);
                sheet.write(i+1,fileNo+3,E);

            elif gas.reaction_type(i) == 4:                               # If falloff reaction
                arr = rxn.low_rate;
                A = arr.pre_exponential_factor;   # A
                b = arr.temperature_exponent;     # b
                E = arr.activation_energy;        # E

                sheet.write(i+1,fileNo+1,A);
                sheet.write(i+1,fileNo+2,b);
                sheet.write(i+1,fileNo+3,E);

            else:                                                         # If PLOG reaction
                arr = rxn.rates;

                for rate in arr:
                    if abs(rate[0]-101325) < dP:
                        dP = abs(rate[0]-101325);
                        pressure = rate[0];

                        A = rate[1].pre_exponential_factor;   # A
                        b = rate[1].temperature_exponent;     # b
                        E = rate[1].activation_energy;        # E

                sheet.write(i+1,fileNo+1,A);
                sheet.write(i+1,fileNo+2,b);
                sheet.write(i+1,fileNo+3,E);

fileExcel = pd.ExcelFile("C:\\Users\\opbir\\Documents\\ML_trainingData.xlsx");   
print(fileExcel.sheet_names)

for sheetName in fileExcel.sheet_names:
    df = pd.read_excel("C:\\Users\\opbir\\Documents\\ML_trainingData.xlsx",sheet_name = sheetName);
    df.fillna(0,inplace=True);

    df.to_excel('C:\\Users\\opbir\\Documents\\ML_trainingData_'+sheetName[:-4]+'.xlsx');
    print(df);

wb.close();
