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

wb = xlsxwriter.Workbook("C:\\Users\\opbir\\Documents\\ML_trainingData_energy.xlsx");
sheet = wb.add_worksheet('Overview');

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

sheet.write(0,0,'energyDFT');
sheet.write(0,1,'atomEnergy');
sheet.write(0,2,'formula');

formulaList = [];
energyDFT = [];
with open("C:\\Users\\opbir\\Documents\\Rad-6_databases2.txt",'r') as f:
    count2 = 1;
    for line in f:
        if "Properties=species:S:1:pos:R:3:uff:R:3" in line:
            energyDFT_indx = line.find("energy");
            energyDFT1 = float(line[energyDFT_indx+7:energyDFT_indx+18]);
            energyDFT.append(energyDFT1);
            sheet.write(count2,0,energyDFT1);

            atomEnergy_indx = line.find("AE");
            if line[atomEnergy_indx+3] == "0":
                atomEnergy = 0;
            else:
                atomEnergy = float(line[atomEnergy_indx+3:atomEnergy_indx+15]);
            sheet.write(count2,1,atomEnergy);

            indx_count = 0;
            smile_indx = line.find("smile");
            while line[smile_indx+7+indx_count] != '"':
                indx_count = indx_count + 1;
            smile = line[smile_indx+7:smile_indx+7+indx_count];

            try:
                mol = Chem.MolFromSmiles(smile);
                formula = CalcMolFormula(mol);
                uniqueAtoms = {atom.GetSymbol() for atom in mol.GetAtoms()};
                sheet.write(count2,2,formula);

                if formula in uniqueSpecies:
                    print(formula)
                    formulaList.append(formula)
            except: 
                pass;

            count2 = count2 + 1;

for filename in os.listdir(directory):
    f = os.path.join(directory,filename)         # full file directory
    if os.path.isfile(f):
        gas = ct.Solution(filename);

        fileNo = 0;
        sheet = wb.add_worksheet(filename);
        sheet.write(0,0,'Reaction');

        count1 = 0;
        for speciesName in uniqueSpecies:
            sheet.write(0,count1+1,speciesName);
            count1 = count1 + 1;

        indx = [i for i, r in enumerate(gas.reactions())];
        for i in indx:
            sheet.write(i+1,fileNo,gas.reaction(i).equation);

            # For each species in a reaction i
            count = 0;            
            for sp in gas.species():
                findSpecies = sp.name;
                if findSpecies in formulaList:
                    location = uniqueSpecies.index(findSpecies);   
                    location_formula = formulaList.index(findSpecies);
                
                    reactantCoeff = gas.reactant_stoich_coeff(findSpecies,i);       # Reactant: -
                    productCoeff = gas.product_stoich_coeff(findSpecies,i);            # Product: +

                    if reactantCoeff != 0:
                        sheet.write(i+1,location+1,energyDFT[location_formula]*96491.5667);             # [J/mol]
                    if productCoeff != 0:
                        sheet.write(i+1,location+1,-1*energyDFT[location_formula]*96491.5667);          # [J/mol]

                count = count + 1;

fileExcel = pd.ExcelFile("C:\\Users\\opbir\\Documents\\ML_trainingData_energy.xlsx");   

for sheetName in fileExcel.sheet_names[1:]:
    df = pd.read_excel("C:\\Users\\opbir\\Documents\\ML_trainingData_energy.xlsx",sheet_name = sheetName);
    df.fillna(0,inplace=True);

    df.to_excel('C:\\Users\\opbir\\Documents\\ML_trainingData_energy_'+sheetName[:-4]+'.xlsx');
    print(df);

wb.close();
