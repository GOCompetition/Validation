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

def run_evaluation(address,files,send_file):
    rawfile = files
    confile = 'case.con'
    inlfile = address+'case.inl'
    ropfile = 'case.rop'
    sol1file = address+files[:-4]+'_scopf_accc_solution1.txt'
    sol2file = address+files[:-4]+'_scopf_accc_solution2.txt'
    summaryfile = address+'summary.csv'
    detailfile = address+'detail.csv'
    #sys.argv = ['python',rawfile, ropfile, confile, inlfile, sol1file, sol2file, summaryfile, detailfile]
    try:
        with silence():
            result1,result2,result3,result4,result5,result6 = evaluation.run(address+rawfile, ropfile, confile, inlfile, sol1file, sol2file, summaryfile, detailfile)
        print ("finished evaluating " + address+files)
        #print(str(result1)+','+str(result2)+','+str(result3)+','+str(result4)+','+str(result5)+','+str(result6)+','+str(100*(1-result2/result1)))#_dict.update({address+files : result })
        send_file.send(str(result1)+','+str(result2)+','+str(result3)+','+str(result4)+','+str(result5)+','+str(result6)+','+str(100*(1-result2/result1)))#_dict.update({address+files : result })
    except IOError:
        print ("Terminate evaluating " + address+files)
        send_file.send('Failed')
            

def run_evaluation_1(address,files):
    rawfile = files
    confile = 'case.con'
    inlfile = address+'case.inl'
    ropfile = 'case.rop'
    sol1file = address+files[:-4]+'_scopf_accc_solution1.txt'
    sol2file = address+files[:-4]+'_scopf_accc_solution2.txt'
    summaryfile = address+'summary.csv'
    detailfile = address+'detail.csv'
    #sys.argv = ['python',rawfile, ropfile, confile, inlfile, sol1file, sol2file, summaryfile, detailfile]
    #with silence():
    result1,result2,result3,result4,result5,result6 = evaluation.run(address+rawfile, ropfile, confile, inlfile, sol1file, sol2file, summaryfile, detailfile)
    print ("finished evaluating " + address+files)
    #send_file.send(str(result1)+','+str(result2)+','+str(result3)+','+str(result4)+','+str(result5)+','+str(result6)+','+str(100*(1-result2/result1)))#_dict.update({address+files : result })


# Main body
if __name__ == "__main__":

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

    # Evaluation 
    if 1:
        return_dict = {}
        procs2 = []
        #p = multiprocessing.Pool(processes = 4) 
        for Counter in Folders:
            all_files = []
            used_raw_files = []

            #Ini_time_Folder = time.clock()

            Location = Case_Location +  Counter + '\\'
            all_files = os.listdir(Location) 
            
            #print "********************************************************"
            print "Identified files in this folder: ", Counter
            for files in all_files:
                if '.raw' in files:
                    #print 'Identified .raw file: ', files
                    used_raw_files.append(files)
                    if 0:
                        # For single processing
                        run_evaluation_1(Location,files)
                    else:
                        # For multiprocessing
                        recv_end,send_end = multiprocessing.Pipe(False)
                        arguments = [Location,files,send_end]
                        p = multiprocessing.Process(target = run_evaluation, args = arguments) # check MAP function
                        procs2.append(p)
                        #print recv_end
                        return_dict.update({Location+files:recv_end})
                        p.start()
     
            while len(procs2)>=3: # number of processes =cores-1
                for p in procs2:
                    if not p.is_alive():
                        procs2.remove(p)
            #print "********************************************************"

        while len(procs2)>0:
            for p in procs2:
                #p.join()
                if not p.is_alive():
                    procs2.remove(p)

    #print return_dict.values()
    #output = [x.recv() for x in return_dict.values()]
    #print output
    # Collecting results only with evaluation multiprocessing
    if 1:
        rslt = open('results.csv','w')
        rslt.write('file,obj,cost,penalty,max_obj_viol,max_nonobj_viol,infeas,100(1-cost/obj)'+'\n')
        output = [x.recv() for x in return_dict.values()]
        scenarios = return_dict.keys()
        #print output
        for irslt in range(len(scenarios)):
            #print irslt, scenarios[irslt]
            #print return_dict[scenarios[irslt]].recv()
            rslt.write(scenarios[irslt]+','+str(output[irslt])+'\n')
            #output = [ x.recv() for x in return_dict.values()]
        #print return_dict, output
        rslt.close()
        #for p in procs:
        #    p.join()