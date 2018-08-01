import sys, os, csv

PSSE_LOCATION = r"C:\Program Files (x86)\PTI\PSSE33\PSSBIN"
sys.path.append(PSSE_LOCATION)
os.environ['PATH'] = os.environ['PATH'] + ';' +  PSSE_LOCATION 

import redirect
redirect.psse2py()
import psspy

import pssarrays
#import matplotlib.pyplot as plt
import csv
import pssexcel
import excelpy
import xlsxwriter
import re
import xlrd
import zipfile
import os
import sys
import math
 
class COMPET_FORM:     
    def __init__(self,cont,solutin1file, solutin2file, vbusno, vbusmag,vbusangle, vgenbusno,vgenid, vgenp, vgenq,vbus_shunt_val, totalgendelta):
        self.cont                                   = cont         
        self.solutin1file                                   = solutin1file
        self.solutin2file                                  = solutin2file
        self.vbusno                               = vbusno
        self.vbusmag                                = vbusmag
        self.vgenbusno                      =vgenbusno
        self.vgenid                      =vgenid
        self.vgenp                              = vgenp
        self.vgenq                              = vgenq
        self.vbusangle                      =vbusangle
        self.vbus_shunt_val             =vbus_shunt_val
        self.totalgendelta              =totalgendelta
        #self.simutimestep                            = simutimestep
        self.Display_Input()
        self.Main_Program()
        
    def Display_Input(self):
        # display input options
        print 'Solution 1 path: ' , self.solutin1file
        print 'Solution 2 path: ' , self.solutin2file
        #print('load_bus_voltage_goal: %f'               % self.load_bus_voltage_goal)
        
    def Main_Program(self):
        
        
        if self.cont == 'InitCase': # this is for solution 1, as it is initial case without contingency
                #put your code to output solution1
            file = open(self.solutin1file,"w") 
            file.write("--bus section\n") 
            file.write("i, v, theta, b\n")
            
            counter=0
            for r in self.vbusno:
                row= str(r)+"," +str(self.vbusmag[counter])+ "," +str(self.vbusangle[counter]) + ","+ str(self.vbus_shunt_val[counter])+"\n"
                file.write(row)
                counter=counter+1
            
            file.write("--generator section\n") 
            file.write("i, uid, p, q\n") 
            counter=0
            for r in self.vgenid:
                row=str(self.vgenbusno[counter])+","+str(r)+"," +str(self.vgenp[counter])+ "," +str(self.vgenq[counter]) + "\n"
                file.write(row)
                counter=counter+1

            file.close()
            
        else:
                # put your code to solution2
            c=1
              
            file = open(self.solutin2file,"a") 
            row="--contingency\n"
            file.write(row)
            file.write("label\n")

            file.write(str(self.cont)+"\n")
            file.write("--bus section\n")
            file.write("i, v, theta, b\n")
            counter=0
            for r in self.vbusno:
                row= str(r)+"," +str(self.vbusmag[counter])+ "," +str(self.vbusangle[counter]) + ","+ str(self.vbus_shunt_val[counter])+ "\n"
                file.write(row)
                counter=counter+1
            file.write("--generator section\n")
            file.write("i, uid, p, q\n")

            counter=0
            for r in self.vgenid:
                row= str(self.vgenbusno[counter])+"," + str(r)+"," +str(self.vgenp[counter])+ "," +str(self.vgenq[counter]) + "\n"
                file.write(row)
                counter=counter+1

    
            file.write("--delta section\n") 
            file.write("delta\n")
            row= str(self.totalgendelta)+"\n"
            file.write(row)
            file.close()  
            
     