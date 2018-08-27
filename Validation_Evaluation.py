# This script was created by Ahmad Tbaileh for any questions contact ahmad.tbaileh@pnnl.gov

from __future__ import with_statement
from contextlib import contextmanager

import os
import sys
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties
import multiprocessing
import random
import copy
import argparse
import evaluation
import csv
from GOValid import GOValid_func
'''
sys_path_PSSE=r'C:\Program Files (x86)\PTI\PSSE33\PSSBIN'  #or where else you find the psspy.pyc
sys.path.append(sys_path_PSSE)
os_path_PSSE=r'C:\Program Files (x86)\PTI\PSSE33\PSSBIN'  # or where else you find the psse.exe
os.environ['PATH'] += ';' + os_path_PSSE
import psspy
throwPsseExceptions = True
import redirect
redirect.psse2py()
import excelpy
import dyntools

psspy.psseinit(1000000)
'''
@contextmanager
def silence(file_object=None):
    """
    Discard stdout (i.e. write to null device) or
    optionally write to given file-like object.
    """
    if file_object is None:
        file_object = open(os.devnull, 'w')

    old_stdout = sys.stdout
    try:
        sys.stdout = file_object
        yield
    finally:
        sys.stdout = old_stdout

# To identify the cases in the current directory
def id_cases(address):
# To create the directories with names of *.sav files
    files = os.listdir(os.curdir)

    sav_files = []

    for file in files:
        if '.raw' in file:
            sav_files.append(file)#[:-4])

    return sav_files


def run_validation_evaluation(address,files):
    rawfile = files
    confile = 'case.con'
    inlfile = 'case.inl'
    monfile = 'All.mon'
    subfile = 'All_SDET.sub'
    with silence():
        GOValid_func(rawfile,confile,inlfile,monfile,subfile,address)
    print ("finished validating " + address+files)
    ropfile = 'case.rop'
    sol1file = address+files[:-4]+'_scopf_accc_solution1.txt'
    sol2file = address+files[:-4]+'_scopf_accc_solution2.txt'
    summaryfile = address+'summary.csv'
    detailfile = address+'detail.csv'
    #sys.argv = ['python',rawfile, ropfile, confile, inlfile, sol1file, sol2file, summaryfile, detailfile]
    with silence():
        evaluation.run(address+rawfile, ropfile, confile, inlfile, sol1file, sol2file, summaryfile, detailfile)
    print ("finished evaluating " + address+files)

# Main body
def main():

    #address = os.getcwd()
    #files=id_cases(address)
    #plotting_function(address,files)

    # Looking into folder and subfolder
    Case_Location = os.getcwd()
    Case_Location = Case_Location + '\\'
    all_files = os.listdir(os.curdir) 
    #print all_files
    used_out_files = []
    Folders = []

    for Counter in all_files:
        #print "here1.2: "
        if not os.path.isfile(Counter):
            #print "here1.3: "
            Folders.append(Counter)
            
            #print "Identified folder: ", Counter

    procs = []
    for Counter in Folders:
        all_files = []
        used_raw_files = []

        #Ini_time_Folder = time.clock()

        Location = Case_Location +  Counter + '\\'
        all_files = os.listdir(Location) 
        
        print "********************************************************"
        print "Identified files in this folder: ", Counter
        for files in all_files:
            if '.raw' in files:
                #print 'Identified .raw file: ', files
                used_raw_files.append(files)
                # For single processing
                run_validation_evaluation(Location,files)
                # For multiprocessing
                #arguments = [Location,files]
                #p = multiprocessing.Process(target = run_validation_evaluation, args = arguments) # check MAP function
                #p.start()
                #procs.append(p)
        print "********************************************************"

if __name__ == "__main__":
    main()