REM ******** RUN PSSE VALIDATION CODE ************
REM ******** argument1: raw case file
REM ******** argument2: contingency file
REM ******** argument3: generator inertial response file
REM ******** argument4: monitored subsystem definition
REM ******** argument5: subsystem definition file

REM ** function      argument1              argument2              argument3              argument4           argument5
python validation.py case14small\case14.raw case14small\case14.con case14small\case14.inl case14small\All.mon case14small\All_SDET.sub
