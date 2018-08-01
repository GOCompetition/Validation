How to Run the Validation Scripts
The whole “Validationcode-release” folder includes the validation code and an input example folder.
The python scripts are scopf_accc_outputsolution.py and COMPET_FORM_class.py
The input example folder is “case14small”
The validation code will run PSS/E preventive security constrained optimal power flow (PSCOPF) for given cases and contingency definition list, try to find a feasible solution for each given case. And then the validation code will run PSS/E AC contingency calculation (ACCC) function for the feasible solution and get the solution 1 and solution 2 files for validation purpose.
Prerequisite to run the validation code: 1) python 2.7 32bit, note PSS/E only support python 2.7 32 bit; 2) PSS/E version 33, the scripts are all tested with PSS/E version 33.7.
To run the validation scripts, you need to create an input subfolder, and the input subfolder should include the following input files:
1) test cases in PTI RAW 33 format, you could put multiple cases (scenarios) in the subfolder and the scripts will automatically scan the subfolder, find all the *.raw files and validate all the cases in raw format in the subfolder. Make sure all the cases in raw format should have the same auxiliary files in (2) to (5).
2) sub file: define the subsystem that ACCC will scan, typically this file remains the same.
3) mon file: define the monitoring elements for the ACCC, typically this file remains the same.
4) con file: define the contingency list
5) inl file: define the generator participation factor
For example, for the input subfolder “case14small”, the following input file are provided:
‘case14.raw’
'All_SDET.sub'
'All.mon'
'case14.con'
'case14.inl'. 
In the python script ‘script is scopf_accc_outputsolution.py’, you should modify line 52 to specify your own input subfolder, and lines 64 to 67 to specify the 4 auxiliary files. Note that you could use the sub file and mon file from the provided “case14small” subfolder, as they are very generic, while you need to 

Once you create your own input subfolder, input files, and modify the python file, to run the validation, just open cmd window, cd to the root “Validationcode_release” path, and then type the following:
python script is scopf_accc_outputsolution.py
The script will generate the solution 1 and solution 2 files in the input subfolder.

To make sure the script runs properly in your computer, you can just open cmd window, cd to the root “Validationcode_release” path, and then type the following without change anything:
python script is scopf_accc_outputsolution.py
This will generate solution1 and solution2 files in the “case14small” folder for the 14-bus case. Ideally, the generated solution1 and solution2 files should be the same as the provided “case14_scopf__solution1_org.txt” and “case14_scopf__solution2_org.txt” files.
The script will output the solution1 and solution 2 file for each raw file in the input subfolder, as well as the script will create a new subfolder with name “inputsubfoldernameyoudefined_scopf”. In this newly created subfolder, the script will save all the feasible solutions from the PSS/E PSCOPF for all the validation cases in the designed input subfolder, with both raw and sav format. In this newly created output subfolder, the script will also create another level of subfolder for each validation case, the ACCC results summary excel file and the post-contingency power flow cases (raw format) for each validation case’s feasible solution will saved in the subfolders.
For questions, please contact:
Renke Huang, renke.huang@pnnl.gov
Ahmad Tbaileh, ahmad.tbaileh@pnnl.gov
Xinda Ke, xinda.ke@pnnl.gov




