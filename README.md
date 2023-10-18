# RNAscope-Data-Analysis
This coding project takes the raw data from RNAscope Image analysis to a clean whole data set 
## Dependencies 
Please ensure you have loaded in the following libraries prior to running the code:
```
    pip install pandas
    pip install os
```
This code was made and checked to run efficiently on a MacOS system on python3.10.
## Execution of Code
1. Ensure that the code file is within the same folder as the same folder/directory of the image file of choice to analyze.
2. You will need to hardcode the absolute file directory of the datasheet for the image in the section where the path is defined.
   ```
    path = "/Users/abigailparakoyi/Desktop/LabWork/20.01.xlsx"
    ```
4. You will need to hardcode the analyzed Excel file name in the def_main() chunk of code. See below example
   ```
   df.to_excel("desired excel name .xlsx")
   ``` 
5. To run the code from the terminal type in the absolute path to the code as seen below
   ```
   /Users/abigailparakoyi/Desktop/LabWork/RNAscope\ Data\ analysis\ Code.py
   ```
   ## Author
   Abigail (Abby) Parakoyi 
