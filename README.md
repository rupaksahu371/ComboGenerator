# ComboGenerator
Tool Description:
Generates unique combinations based on given input data set and give output in excel workbook. It can generate upto 1M combination based on your given input data products. To achieve this we are using pandas, itertools, tqdm and math libraries.

Input Changes: You can change inputs in InputValues.xlsx file.
1. To add more columns just or change column name you can directly write on row 1,
2. Ensure your data set has unique data value in fields for desire output,
3. Remove A column while using ComboGenerator Tool,

Monitoring: You can monitor your output result before running in excel.
1. For output combination you can check Summary worksheet,
2. Drag the 3rd and 4th row formula in case you have more than 10 columns in Input Master worksheet,
3. Select "Max_column + 4" Range in B7 (Combination) end cell,

After changing input file save and run the python script for output. It will be give output on same location of InputMaster.xlsx or "Combination Generator.py" As file name : 'CombinedData_'+ current_time +'.xlsx'

Enjoy! 🙌
