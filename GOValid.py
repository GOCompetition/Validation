import sys, os, csv

PSSE_LOCATION = r"C:\Program Files (x86)\PTI\PSSE33\PSSBIN"
sys.path.append(PSSE_LOCATION)
os.environ['PATH'] = os.environ['PATH'] + ';' +  PSSE_LOCATION 
from COMPET_FORM_class import COMPET_FORM
import redirect
redirect.psse2py()
import psspy
import re

import pssarrays
#import matplotlib.pyplot as plt
import csv
import pssexcel
import excelpy
import xlsxwriter
import re
import xlrd
import zipfile
import shutil

# To identify the cases in the current directory
def id_cases(address):

# To create the directories with names of *.raw files
    files = os.listdir(address)

    sav_files = ['']

    for file in files:
        if '.raw' in file:
        #if '.sav' in file:    
            sav_files.append(file[:-4])

    return sav_files

### ------------------- input output data section --------------------------------------------------------------------
# testcasecur is the input data folder name
# the cases inside this folder should be PTI RAW 33 format, and the following supporting files are also needed:
# sub file: define the subsystem ACCC will run
# mon file: define the monitoring elements for the ACCC
# con file: define the contingency list
# inl file: define the generator participation factor

# you need to change the name of the testcasecur at line 52 to run validation for other cases
# please modify the names of the supporting files from line 64 to 67 accordingly


def GOValid_func(rawfile,confile,inlfile,monfile,subfile):
    # This is the 'root' directory name for a set of cases and supporting files
    # please modify it accordingly
    case = str(rawfile)[:-4]
    testcasecur = case
    # defining the supporting files, please modify them accordingly
    fileSub = str(subfile)
    fileMon = str(monfile)
    fileCon = str(confile)
    fileINL = str(inlfile)

    address = os.getcwd()
    cur_dir = os.getcwd()
    scopfaddress = os.getcwd() + '\\' + testcasecur + '_scopf' # this is the output data folder

    if not os.path.isdir(scopfaddress):
        os.makedirs(scopfaddress)

    
    ### ------------------- input output data section ends -------------------------------------------------------------------

    # Options to performe the ACCC simulation
    tol = 0.5 

    options = [0, 0, 0, 1, 1, 0, 0, 0, 1, 3, 1] 

    subsystem = 'ALL'

    # The instructions to export to excel
    string = ['s', 'e', 'v', 'b']
    colabel = ''
    stype = 'contingency'
    busmsm=0.5
    sysmsm=5.0
    rating = 'a'
    namesplit = True
    sheet = ''
    overwritesheet = True
    show = True
    baseflowvio = True
    basevoltvio = True
    overwritesheet = True
    show = False
    ratecon = 'b' # changed from a to b
    flowlimit = 100.0
    flowchange = 1.0
    voltchange = 0.05

    # the following csv and excel file record whether the SCOPF for each case is successful
    csv_scopf_file = scopfaddress+'\\'+ 'SCopfresults_0.csv'
    csv_scopf_obj = open(csv_scopf_file, 'w', buffering=1)
    csv_scopf_writer = csv.writer(csv_scopf_obj)
    listtmp = ['Case', 'err_code']
    csv_scopf_writer.writerow(listtmp)

    workbookscopf = xlsxwriter.Workbook(scopfaddress+'\\'+ 'SCopfresults_0.xlsx')
    xscopf = workbookscopf.add_worksheet()
    Row = 1
    xscopf.write("A" + str(Row),'Case')   
    xscopf.write("B" + str(Row),'err_code')
    Row = Row + 1

    psspy.psseinit(100000)
    _i=psspy.getdefaultint()
    _f=psspy.getdefaultreal()
    _s=psspy.getdefaultchar()

    print ('start SCOPF analysis for case: ' + case)
    
    psspy.psseinit(1000000)

    #psspy.case(case) #this is for sav file
    psspy.read(0,case+'.raw') #this is for raw file
    psspy.fnsl([0,0,0,1,1,0,0,0])

    # parse the con file to make sure there is no swing bus generator contingency, 
    # and construct a dictionary for the generator contingencies
    print ('------------------start checking swing bus in con file -------------------')
    ierr, iarray = psspy.abusint(-1, 1, 'NUMBER')
    vtmpbusno = iarray[0]   
    ierr, iarray = psspy.abusint(-1, 1, 'TYPE')
    vtmpbustype = iarray[0] 
    swingbus_str = []
    swingbus_tmp = []
    for ibus in range(0, len(vtmpbusno)):
        if vtmpbustype[ibus] == 3:
            swingbus_str.append(str(vtmpbusno[ibus]))
            swingbus_tmp.append(vtmpbusno[ibus])
    
    fcon = open (fileCon)
    conlines = fcon.readlines()
    cont_con_array = []
    cont_gen_dict = {}
    cont_genbus_array = []
    swingbusincont = False
    for oneline in conlines:
        #partxt = re.split(r'[\s]',oneline)
        partxt = oneline.split()
        #print partxt
        
        if partxt[0] == 'CONTINGENCY':
            cont_tag = partxt[1]
            cont_con_array.append(cont_tag)
            
        if (partxt[0] == 'REMOVE' and partxt[1] == 'UNIT') or (partxt[0] == 'REMOVE' and partxt[1] == 'MACHINE'):  
            
            valtmp = (partxt[5], partxt[2])
            cont_gen_dict.update({cont_tag:valtmp})
            cont_genbus_array.append(partxt[5])
            
            if partxt[5] in swingbus_str:
                swingbusincont = True
    
    fcon.close()
    
    print ('-------swingbus_str: ')
    print (swingbus_str)
    print ('------swingbusincont:  ' )
    print (swingbusincont)
    
    # check if swing bus in contingency:
    if swingbusincont:
        #sort in-service generators
        ierr, iarray = psspy.amachint(-1, 1, 'NUMBER')
        vtmpgenbusno = iarray[0] 
        ierr, rarray = psspy.amachreal(-1, 1, 'PGEN')
        vtmpgenpgen = rarray[0]
        gen_tmp_info = zip(vtmpgenbusno, vtmpgenpgen) 
        gen_tmp_sorted = sorted(gen_tmp_info, key=lambda item:item[1], reverse=True)   
        
        newswingbus = -1
        for igen in range(0, len(gen_tmp_sorted)):
            if  str(gen_tmp_sorted[igen][0]) not in cont_genbus_array:
                newswingbus = gen_tmp_sorted[igen][0]
                break
                
        if newswingbus != -1:
            print ('!!!!!!!!!!!--------new swing bus find, is bus: ' + str(newswingbus) + '   ----------------!!!!!')   
            
            # change swing bus:
            psspy.bus_chng_3(newswingbus,[3,_i,_i,_i],[_f,_f,_f,_f,_f,_f,_f],_s)  # change swing to type 3   
            for ibus in range(0, len(swingbus_tmp)):
                psspy.bus_chng_3(swingbus_tmp[ibus],[2,_i,_i,_i],[_f,_f,_f,_f,_f,_f,_f],_s)  # change swing to type 2
            
        psspy.fnsl([0,0,0,1,1,0,0,0])    
        #psspy.save(address + '\\' + caseX + '_swigchng.sav')      
    
    print ('------------------finish checking swing bus in con file -------------------')
    
    # scan the inl file to make sure the Pmax and Pmin element for each record is not 0.0, or equal 
    # creat new inl file if there is any Pmax = Pmin
    # PSS/E ACCC function will not dispatch generators if Pmax = Pmin  
    
    print ('------------------start checking Pmax Pmin in inl file -------------------')
    
    ierr, iarray = psspy.amachint(-1, 4, 'NUMBER')
    vgenbusnotmp = iarray[0] # this array has all the generator's bus number, including both in-service and out-service
    ierr, iarray = psspy.amachint(-1, 4, 'STATUS')
    vgenstatustmp = iarray[0] # this array has all the generator's status: in-service (1) and out-service (0)
    ierr, carray = psspy.amachchar(-1, 4, 'ID')
    vgenidtmp = carray[0] # this array has all the generator's ID, string
    ierr, rarray = psspy.amachreal(-1, 4, 'PMAX')
    vgenpmaxtmp = rarray[0] # this array has all the generator's Active power output P, MW
    ierr, rarray = psspy.amachreal(-1, 4, 'PMIN')
    vgenpmintmp = rarray[0] # this array has all the generator's Reactive power output Q, MVar
    
    geninfotmp = zip(vgenpmaxtmp, vgenpmintmp)
    genbusdicttmp = {}
    for igentmp in range(0, len(vgenbusnotmp)):
        genbusdicttmp.update({vgenbusnotmp[igentmp]: geninfotmp[igentmp]}) 
                
    #shutil.copyfile (fileINL, fileINL[:-4]+'_org.inl') #first keep a copy of the original inl file
    fileINLorg = fileINL[:-4]+'_org.inl'

    if os.path.exists(fileINLorg):
        os.remove (fileINLorg)
        os.rename (fileINL, fileINLorg)
    else:
        os.rename (fileINL, fileINLorg)
    
    finlorg = open (fileINLorg)
    inllines = finlorg.readlines()
    
    finldst = open(fileINL, 'w')

    for oneline in inllines:
        if oneline.split()[0] == '0':
            break
        
        partxt = oneline.split(',')
        igenbustmp = int(partxt[0])
        if igenbustmp in genbusdicttmp:
            if abs(float(partxt[3]) - 0.0 )<0.000001 and abs(float(partxt[4]) - 0.0 )<0.000001:
                
                str_pmax = "%6.3f" %(genbusdicttmp[igenbustmp][0]/100.0)
                str_pmin = "%6.3f" %(genbusdicttmp[igenbustmp][1]/100.0)
                finldst.write(' '+ partxt[0].strip() + ',   ' + partxt[1].strip() + ',  ' + partxt[2].strip() + ',  ' + str_pmax + ',  ' + str_pmin + ',  ' + partxt[5].strip()  + ',  ' + partxt[6].strip() + '\n')
            else:
                finldst.write(oneline)
        else:
            finldst.write(oneline)
   
    finlorg.close()
    
    finldst.write(str(0) )
    finldst.close()
    
    print ('------------------finish checking Pmax Pmin in inl file -------------------')
    
    # prepare the participation factor file for ACCC and SCOPF
    scopfdfx = scopfaddress+'\\'+ case + '.dfx'
    #accout = address + '\\' + caseX + '\\' + caseX + '.acc'
    #Progress = address + '\\' + caseX + '\\' + 'Progress_' + caseX + '.txt'
    #logFile = file(Progress, "a")
    #sys.stdout = logFile
    psspy.dfax([1, 1], fileSub, fileMon, fileCon, scopfdfx)
    psspy.solution_parameters_4(intagar=[_i,150,_i])

    '''
    # may need to change the normal volt max and min here
    
    ierr, iarray = psspy.abusint(-1, 2, 'NUMBER')
    vbusno = iarray[0]
    for ibus in range(0, len(vbusno)):
        psspy.bus_chng_3(vbusno[ibus], realar4 = 1.05, realar5 = 0.95)  
    
    
    psspy.fnsl([0,0,0,1,1,0,0,0]) 
    
    psspy.rawd_2(0,1,[1,1,1,0,0,0,0],0, address + '\\' + caseX +'lim.raw')
    '''
    
    # RUN SCOPF from PSS/E
    print ('------------------start  SCOPF for case:' + case + '  ------------------')
    
    ierr = psspy.pscopf_2([0,0,0,0,1,0,1,0,0,0,0,0,1,0,3,1,2,1,2,30,2,1,1,0,0,0,0,1],
                   [ 0.5, 100.0, 98.0, 0.02, 0.1, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0],
                   [r"""ALL""",r"""ALL""",r"""ALL""",r"""ALL""",r"""ALL""",r"""ALL""",r"""ALL"""],
                   scopfdfx,fileINL, "")
    
    # write the flag whether the SCOPF is successful or not
    xscopf.write("A" + str(Row), case)  
    xscopf.write("B" + str(Row), ierr)        
    Row = Row + 1 
                    
    listtmp2 = [case, ierr]
    csv_scopf_writer.writerow(listtmp2)
    
    # save case
    savecase = scopfaddress + '\\' + case + '_scopf.sav'
    psspy.save(savecase)
    psspy.rawd_2(0,1,[1,1,1,0,0,0,0],0,scopfaddress + '\\' + case +'_scopf.raw')   
    
    print ('------------------finish SCOPF for case:' + case + '  ------------------')
    
    #---------------------form the base case gen dictionary-------------------------------
    ierr, iarray = psspy.amachint(-1, 4, 'NUMBER')
    vbasecasegenbusno = iarray[0] # this array has all the generator's bus number, including both in-service and out-service
    ierr, iarray = psspy.amachint(-1, 4, 'STATUS')
    vbasecasegenstatus = iarray[0] # this array has all the generator's status: in-service (1) and out-service (0)
    ierr, carray = psspy.amachchar(-1, 4, 'ID')
    vbasecasegenid = carray[0] # this array has all the generator's ID, string
    ierr, rarray = psspy.amachreal(-1, 4, 'PGEN')
    vbasecasegenp = rarray[0] # this array has all the generator's Active power output P, MW
    ierr, rarray = psspy.amachreal(-1, 4, 'QGEN')
    vbasecasegenq = rarray[0] # this array has all the generator's Reactive power output Q, MVar
    
    ierr, iarray = psspy.abusint(-1, 1, 'NUMBER')
    vbasebusno = iarray[0]   # this array has all the bus number
    ierr, iarray = psspy.abusint(-1, 1, 'TYPE')
    vbasebustype = iarray[0]   # this array has all the bus number
    
    swingbus = []
    for ibustmp in range(0, len(vbasebusno)):
        if vbasebustype[ibustmp] == 3:
            swingbus.append(vbasebusno[ibustmp])
    
    basecase_gen_dict = {}
    swing_gen = []
    
    for igentmp in range(0, len(vbasecasegenbusno)):
        genbuskeytmp = str(vbasecasegenbusno[igentmp])+'-'+vbasecasegenid[igentmp].strip()
        genptmp = vbasecasegenp[igentmp] * vbasecasegenstatus[igentmp]
        basecase_gen_dict.update({genbuskeytmp:genptmp})
        
        # here we also need to find the swing generators
        if vbasecasegenbusno[igentmp] in swingbus:
            swing_gen.append(genbuskeytmp)
    
    #---------------------form the base case gen dictionary finished here-------------------------------
    
    # run ACCC for the new case from SCOPF
    psspy.case(savecase)
    #psspy.read(0,case)
    psspy.fnsl([0,0,0,1,1,0,0,0])
    #psspy.rawd_2(0,1,[1,1,1,0,0,0,0],0, address + '\\' + caseX +'.raw')
    
    # creat sub-folder to store all the ACCC results
    case = case + '_scopf_accc'
    if not os.path.isdir(scopfaddress + '\\' + case ):
        os.makedirs(scopfaddress + '\\' + case )
    
    acccdfx = scopfaddress  + '\\' + case+ '\\' + case + '.dfx'
    accout = scopfaddress + '\\' + case+ '\\' + case + '.acc'
    Zip = scopfaddress + '\\' + case+ '\\' + case + '.zip'
    #Progress = address + '\\' + caseX + '\\' + 'Progress_' + caseX + '.txt'
    #logFile = file(Progress, "a")
    #sys.stdout = logFile
    psspy.dfax([1, 1], fileSub, fileMon, fileCon, acccdfx)
    psspy.solution_parameters_4(intagar=[_i,150,_i])
    
    # run ACCC
    psspy.accc_with_dsp_3(tol ,options,'ALL', acccdfx, accout, "", fileINL,Zip)
    xlsfile = scopfaddress+ '\\' + case+ '\\' + case + '.xlsx'
    pssexcel.accc(accout, string, colabel, stype, busmsm, sysmsm, rating,
                  namesplit, xlsfile, sheet, overwritesheet, show, ratecon,baseflowvio, basevoltvio, flowlimit, flowchange, voltchange)
    excelfile = excelpy.workbook(xlsfile)
    excelfile.close()
    
    # Saving post-contingency cases
    #pywin.debugger.set_trace()
    archive = zipfile.ZipFile(Zip, 'r')
    ziplist = archive.namelist()
    isvfiles = []
    for file in ziplist:
        if '.isv' in file:
            isvfiles.append(file[:-4])
        if '.sav' in file:
            savefile = file
    #pywin.debugger.set_trace()
    # Get contingency names from excel sheet
    workbook_accc = xlrd.open_workbook(xlsfile)
    sheet_bf = workbook_accc.sheet_by_name('Contingency Events')
    
    cases_delta_dict = {}
    
    for isvfile in isvfiles:
        #print isvfile
        if isvfile!='InitCase':
            contno = re.findall('\d+',isvfile)#[int(s) for s in isvfile if s.isdigit()]
            #print contno
            row = int(contno[0])*2+1
            #print row
            cont = sheet_bf.cell_value(row,0)
            
            # remove cont contingency in the cont_con_array, to check at the end whether all the contingencies are processed
            if cont in cont_con_array:
                cont_con_array.remove(cont)
                
        else:
            cont = 'InitCase'
        ierr = psspy.getcontingencysavedcase(Zip, isvfile)  
        
        # extract data for solution 1 and solution 2
        # bus section
        ierr, iarray = psspy.abusint(-1, 1, 'NUMBER')
        vbusno = iarray[0]   # this array has all the bus number
        print "type:", type(vbusno)
        ierr, rarray = psspy. abusreal(-1, 1, 'PU')  
        vbusmag = rarray[0]  # this array has all the bus voltage magnitude
        ierr, rarray = psspy. abusreal(-1, 1, 'ANGLE')  
        vbusangle = rarray[0]      # this array has all the bus voltage angle, in ardians
        
        # generator section
        ierr, iarray = psspy.amachint(-1, 4, 'NUMBER')
        vgenbusno = iarray[0] # this array has all the generator's bus number, including both in-service and out-service
        ierr, iarray = psspy.amachint(-1, 4, 'STATUS')
        vgenstatus = iarray[0] # this array has all the generator's status: in-service (1) and out-service (0)
        ierr, carray = psspy.amachchar(-1, 4, 'ID')
        vgenid = carray[0] # this array has all the generator's ID, string
        ierr, rarray = psspy.amachreal(-1, 4, 'PGEN')
        vgenp = map(lambda (a,b):a*b,zip(vgenstatus,rarray[0] ))    # this array has all the generator's Active power output P, MW
        ierr, rarray = psspy.amachreal(-1, 4, 'QGEN')
        vgenq = map(lambda (a,b):a*b,zip(vgenstatus,rarray[0] )) # this array has all the generator's Reactive power output Q, MVar

        # switched shunts section
        ierr, iarray = psspy.aswshint(-1, 4, 'NUMBER')
        swshuntbusno = iarray[0] # this array has all the switched shunts bus number
    
        ierr, iarray = psspy.aswshint(-1, 4, 'STATUS')
        swshuntstatus = iarray[0] # this array has all the switched shunts status
    
        ierr, rarray = psspy.aswshreal(-1, 4, 'BSWNOM')
        swshunt_inival = rarray[0] # this array has all the switched shunts status
        
        # form the switched shunt dictionary
        shunt_dict = {}
        vbus_shunt_val = []
        for ibustmp in range(0, len(swshuntbusno)):
            shuntkeytmp = swshuntbusno[ibustmp]
            shuntval = swshuntstatus[ibustmp] * swshunt_inival[ibustmp]
            shunt_dict.update({shuntkeytmp:shuntval})
            
        # form the switched shunt values (b) for each bus
        for ibustmp in range(0, len(vbusno)):
            busnotmp = vbusno[ibustmp]
            if busnotmp in shunt_dict.keys():
                vbus_shunt_val.append(shunt_dict[busnotmp])
            else:
                vbus_shunt_val.append(0.0)
        
        # start compute the delta value for the contingency
        vgen_delta_dict= {}
        totalgendelta = 0.0
        
        for igentmp in range(0, len(vgenbusno)):
            genbuskeytmp = str(vgenbusno[igentmp])+'-'+vgenid[igentmp].strip()
            genptmp = vgenp[igentmp] * vgenstatus[igentmp]
            
            basegenp = basecase_gen_dict[genbuskeytmp]
            deltatmp = genptmp - basegenp
            
            '''
            # here we also need to excluded the swing generators
            if genbuskeytmp in swing_gen:
                deltatmp = 0.0
                print('!!!!testout-----------------Swing Gen is ' + genbuskeytmp + ': ' + str(deltatmp))
            '''
                
            vgen_delta_dict.update({genbuskeytmp:deltatmp})                    
            tmpstr = 'GEN-'+genbuskeytmp
                            
            if cont in cont_gen_dict.keys():
                
                contgen_info = cont_gen_dict[cont]   
                #print ('contgen_info')
                #print contgen_info
                
                if str(vgenbusno[igentmp]) == contgen_info[0] and vgenid[igentmp].strip() == contgen_info[1] :
                    deltatmp = 0.0
                    print('!!!!testout-----------------Outrage Gen is ' + tmpstr + ': ' + str(deltatmp))
                    
            totalgendelta = totalgendelta + deltatmp        
        
        #totalgendelta is the delta value for the case    
        cases_delta_dict.update ({cont:totalgendelta})  
        
        if cont == 'InitCase':
            vbusno_basecse      = vbusno
            vbusmag_basecse     = vbusmag
            vbusangle_basecse   = vbusangle
            vgenbusno_basecse   = vgenbusno
            vgenid_basecse      = vgenid
            vgenp_basecse       = vgenp
            vgenq_basecse       = vgenq
            vbus_shunt_val_basecse = vbus_shunt_val
            totalgendelta_basecse = 0.0

        # you can use the special string cont as the indicator of which contingency is for the solution 2:
        # if you want to write the solution 2 of each contingency in separate files, the separate file for solution 2 should be:
        # solutin2file = address + '\\' + caseX +'\\'+caseX +"_" + str(cont)+'_solution2.txt'
        
        ############ below we start to write solution file for compeitiont format
        solutin1file = address + '\\' +case  + '_solution1.txt'
        solutin2file = address + '\\' +case + '_solution2.txt'

        System = COMPET_FORM(cont,solutin1file, solutin2file, vbusno, vbusmag,vbusangle,vgenbusno,vgenid, vgenp, vgenq, vbus_shunt_val,totalgendelta)
        
        # save each post-contingency case in both raw and sav format
        # psspy.save(scopfaddress + '\\' + caseX +'\\'+caseX +"_" + str(cont)+'.sav')
        psspy.rawd_2(0,1,[1,1,1,0,0,0,0],0,scopfaddress + '\\' + case +"_" + str(cont)+'.raw')
    
    # if cont_con_array is not empty, which means ACCC ignores some contingencies, 
    # output solution1 as solution2 for these contingencies 
    if cont_con_array:
        print ('!!!!!!!!!!!!------------cont_con_array is not empty!--------------')
        for icon in range(0, len(cont_con_array)):
            conttmp = cont_con_array[icon]
            System = COMPET_FORM(conttmp,solutin1file, solutin2file, vbusno_basecse, vbusmag_basecse,vbusangle_basecse, \
                                 vgenbusno_basecse,vgenid_basecse, vgenp_basecse, vgenq_basecse, vbus_shunt_val_basecse,totalgendelta_basecse)
    
    # for each case, need to clear PSS/E memory to start a new one 
    psspy.pssehalt_2()
    
    csv_scopf_obj.close()  
    workbookscopf.close()   

    if os.path.exists(fileINL[:-4]+'_mod.inl'):
        os.remove (fileINL[:-4]+'_mod.inl')
        os.rename (fileINL, fileINL[:-4]+'_mod.inl')
    else:
        os.rename (fileINL, fileINL[:-4]+'_mod.inl')
        
    os.rename (fileINLorg, fileINL)

    for casetmp in cases_delta_dict.keys():
        print(casetmp  + '  delta:  ' + str(cases_delta_dict[casetmp]))
            
    print ' +++++++++++++++done'
            

'''
rawfile = 'case14.raw'
confile = 'case14.con'
inlfile = 'case14.inl'
monfile = 'All.mon'
subfile = 'All_SDET.sub'
GOValid(rawfile,confile,inlfile,monfile,subfile)

GOValid('case14.raw','case14.con','case14.inl','All.mon','All_SDET.sub')
'''