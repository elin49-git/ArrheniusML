# ArrheniusML Contents

**Note**: The directory of the files used in the code/scripts will not be the same between workstations. 

**Code/Scripts**:
1. **Arrhenius ML_final.ipynb**: Jupyter notebook code containing all the machine learning steps.
2. **elemental.py**: calculates the number of C, H, O, N, and AR atoms for each molecule.
3. **moleFractions.py**: calculates the mole fraction of each species in a reaction.
4. **rad6_atomEnergy.py**: extracts the atomization energy from the rad-6 database.
5. **rad6_molecules.py**: extracts the moleuclar properties for each relevant species from the rad-6 database.
6. **reactionData.py**: extracts the thermophysical properties for each species and reaction from CANTERA. 

**Datasets**:
1. **ML Training Data**: contains most of the datasets used for machine learning. Training data were pre-processed using above codes.
2. **ML_trainingData.xlsx**: dataset pre-processed by reactionData.py.
3. **ML_trainingData_element_df.xlsx**: dataset pre-processed by elemental.py.
