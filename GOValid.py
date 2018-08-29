# 0816 - includes correction to trigger dispatch in case of line loss
# 0817 - uses the slack generator droop value to improve dispatch value
# 0818 - changes shunt control setting to continuous, resets low and high voltage to current bus voltage
# 0821 - change voltage set point for generators at Qmax or Qmin
# 0822 - imposing Pmin and Pmax from case, Mbase = 100
#       - new way to calculate delta
# 0823 - fixing the second round of delta calculation after INLF
#       - use deltamean rather than genloss in correction equation
#       - impose Bmin Bmax Qmin Qmax on extracted values
#       - change remote bus to 0 (self) for all generators
#       - change voltage set point after SCOPF
# 0824 - including eps tolerance for Bmin and Bmax
# 0825 - inl modified copy saved in subfolder
#       - remove existing solution files
import sys, os, csv

PSSE_LOCATION = r"C:\Program Files (x86)\PTI\PSSE33\PSSBIN"
sys.path.append(PSSE_LOCATION)
os.environ['PATH'] = os.environ['PATH'] + ';' +  PSSE_LOCATION 
from COMPET_FORM_class import COMPET_FORM
import redirect
redirect.psse2py()
import psspy
import re
import numpy
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


def GOValid_func(rawfile,confile,inlfile,monfile,subfile,address):
    # This is the 'root' directory name for a set of cases and supporting files
    # please modify it accordingly
    case = str(rawfile)[:-4]
    testcasecur = case
    # defining the supporting files, please modify them accordingly
    fileSub = str(subfile)
    fileMon = str(monfile)
    fileCon = str(confile)
    fileINL = str(inlfile)

    #address = os.getcwd()
    cur_dir = os.getcwd()
    scopfaddress = address + testcasecur + '_scopf' # this is the output data folder

    if not os.path.isdir(scopfaddress):
        os.makedirs(scopfaddress)

    
    ### ------------------- input output data section ends -------------------------------------------------------------------
    eps = 1e-8

    # Options to performe the ACCC simulation
    tol = 0.5 
    
    options = [0, 0, 0, 1, 1, 0, 0, 0, 1, 4, 1] 

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
    ratecon = 'c' # changed from a to c
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
    psspy.read(0,address+case+'.raw') #this is for raw file
    psspy.fnsl([0,0,0,1,1,0,0,0])


    # Change remote voltage set point
    if 1:
        print ('------------------ change remote bus for all generators to self (0) ---------')
        ierr, iarray = psspy.amachint(-1, 1, 'NUMBER')
        MachBus = iarray[0] 
        for imach in range(0,len(MachBus)):
            ierr = psspy.plant_chng(MachBus[imach], intgar1=0)
            
        print ('------------------ finished change remote bus for all generators to self (0) ---------')

    if 0:
        print ('------------------ Change all switched shunts to continuous control mode ---------')
        ierr, iarray = psspy.aswshint(-1, 4, 'NUMBER')
        ShuntBus = iarray[0]
        for ishunt in range(0,len(ShuntBus)):
            ierr, Vpu = psspy.busdat(ShuntBus[ishunt] ,'PU')
            ierr = psspy.switched_shunt_chng_3(ShuntBus[ishunt], intgar9=2, realar9=Vpu, realar10=Vpu)
        print ('------------------ finished changing all switched shunts to continuous control mode ---------')

    ierr =  psspy.fnsl([0,0,0,1,1,0,0,0])
    
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
    vgenpmaxtmp = rarray[0] # this array has all the generator's Pmax, MW
    ierr, rarray = psspy.amachreal(-1, 4, 'PMIN')
    vgenpmintmp = rarray[0] # this array has all the generator's Pmin, MW
    
    geninfotmp = zip(vgenpmaxtmp, vgenpmintmp)
    genbusdicttmp = {}
    gendroop = {}
    for igentmp in range(0, len(vgenbusnotmp)):
        genbuskeytmp = str(vgenbusnotmp[igentmp])+'-'+vgenidtmp[igentmp].strip()
        genbusdicttmp.update({genbuskeytmp: geninfotmp[igentmp]})
        # Set Mbase to 100 for all machines
        ierr = psspy.machine_chng_2(vgenbusnotmp[igentmp], vgenidtmp[igentmp],  realar7 = 100.0)

    # We will just a store copy of the modified INL file in each subfolder (for multiprocessing)      
    #shutil.copyfile (fileINL, fileINL[:-4]+'_org.inl') #first keep a copy of the original inl file
    #fileINLorg = fileINL[:-4]+'_org.inl'

    if os.path.exists(address+fileINL):
        os.remove (address+fileINL)
        #os.rename (fileINL, fileINLorg)
    #else:
    #    os.rename (fileINL, fileINLorg)
    
    finlorg = open (fileINL)
    inllines = finlorg.readlines()

    finldst = open(address+fileINL, 'w')

    for oneline in inllines:
        if oneline.split()[0] == '0':
            break
        
        partxt = oneline.split(',')
        igenbustmp = int(partxt[0])
        igenidtmp = str(partxt[1].strip())
        genbuskeytmp = str(igenbustmp)+'-'+igenidtmp
        if genbuskeytmp in genbusdicttmp.keys():
            # Storing the generators droop for delta calculation
            #genbuskeytmp = str(igenbustmp)+'-'+igenidtmp
            gendroop.update({genbuskeytmp: float(partxt[5].strip())})
            # checking pmax and pmin values in inl file
			#if abs(float(partxt[3]) - 0.0 )<0.000001 and abs(float(partxt[4]) - 0.0 )<0.000001:    
            if 1: # Impose Pmin and Pmax from case 
                str_pmax = "%6.3f" %(genbusdicttmp[genbuskeytmp][0]/100.0)
                str_pmin = "%6.3f" %(genbusdicttmp[genbuskeytmp][1]/100.0)
                finldst.write(' '+ partxt[0].strip() + ',   ' + partxt[1].strip() + ',  ' + partxt[2].strip() + ',  ' + str_pmax + ',  ' + str_pmin + ',  ' + partxt[5].strip()  + ',  ' + partxt[6].strip() + '\n')
            else:
                finldst.write(oneline)
        else:
            finldst.write(oneline)
   
    finlorg.close()
    
    finldst.write(str(0) )
    finldst.close()
    fileINL = address+fileINL
    
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
    
    ierr = psspy.pscopf_2([0,0,0,0,1,0,1,0,0,0,0,0,1,0,4,1,3,1,2,30,2,1,1,0,0,0,0,1],
                   [ 0.5, 100.0, 98.0, 0.02, 0.1, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0],
                   [r"""ALL""",r"""ALL""",r"""ALL""",r"""ALL""",r"""ALL""",r"""ALL""",r"""ALL"""],
                   scopfdfx,fileINL, "")
    
    # write the flag whether the SCOPF is successful or not
    xscopf.write("A" + str(Row), case)  
    xscopf.write("B" + str(Row), ierr)        
    Row = Row + 1 
                    
    listtmp2 = [case, ierr]
    csv_scopf_writer.writerow(listtmp2)
    

    # Change voltage set point for generators at Qmin or Qmax to current bus voltage
    if 1:
        print ('------------------ change voltage set point for generators at Qmax or Qmin ---------')
        ierr, iarray = psspy.amachint(-1, 1, 'NUMBER')
        MachBus = iarray[0] 
        #ierr, carray = psspy.amachchar(-1, 1, 'ID')
        #MachID = carray[0]
        ierr, rarray = psspy.amachreal(-1, 1, ['QGEN','QMAX','QMIN'])
        Qgen = rarray[0]
        Qmax = rarray[1]
        Qmin = rarray[2]

        for imach in range(0,len(MachBus)):
            if Qgen[imach]==Qmax[imach] or Qgen[imach]==Qmin[imach]:
                ierr, Vpu = psspy.busdat(MachBus[imach] ,'PU')
                ierr = psspy.plant_chng(MachBus[imach], realar1 = Vpu)
            
        print ('------------------ finished change voltage set point for generators at Qmax or Qmin ---------')

    # save case
    savecase = scopfaddress + '\\' + case + '_scopf.sav'
    psspy.save(savecase)
    psspy.rawd_2(0,1,[1,1,1,0,0,0,0],0,scopfaddress + '\\' + case +'_scopf.raw')   
    
    print ('------------------finish SCOPF for case:' + case + '  ------------------')
    
	
    # run ACCC for the new case from SCOPF
    psspy.case(savecase)
    #psspy.read(0,case)
    psspy.fnsl([0,0,0,1,1,0,0,0])
    #psspy.rawd_2(0,1,[1,1,1,0,0,0,0],0, address + '\\' + caseX +'.raw')
    
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
    # getting the machines bus voltages
    #ierr = psspy.bsys(11, 0, [0.2,999.0], 0, [], len(vbasecasegenbusno), vbasecasegenbusno,0, [], 0, [])
    #ierr, rarray = psspy. abusreal(11, 1, 'PU')
    #vbasecasegenbusvpu = rarray[0]
    
    ierr, iarray = psspy.abusint(-1, 1, 'NUMBER')
    vbasebusno = iarray[0]   # this array has all the bus number
    ierr, iarray = psspy.abusint(-1, 1, 'TYPE')
    vbasebustype = iarray[0]   # this array has all the bus number
    
    swingbus = []
    basebusvpu_dict = {}
    for ibustmp in range(0, len(vbasebusno)):
        if vbasebusno[ibustmp] not in basebusvpu_dict.keys(): 
            ierr, vpu = psspy.busdat(vbasebusno[ibustmp],'PU')
            basebusvpu_dict.update({vbasebusno[ibustmp]:vpu})
        if vbasebustype[ibustmp] == 3:
            swingbus.append(vbasebusno[ibustmp])
    
    basecase_gen_dict = {}
    basecase_gen_dict_stat = {}
    swing_gen = []
    
    for igentmp in range(0, len(vbasecasegenbusno)):
        genbuskeytmp = str(vbasecasegenbusno[igentmp])+'-'+vbasecasegenid[igentmp].strip()
        genptmp = vbasecasegenp[igentmp] * vbasecasegenstatus[igentmp]
        basecase_gen_dict.update({genbuskeytmp:genptmp})
        basecase_gen_dict_stat.update({genbuskeytmp:vbasecasegenstatus[igentmp]})
        
        # here we also need to find the swing generators
        if vbasecasegenbusno[igentmp] in swingbus:
            swing_gen.append(genbuskeytmp)
    
    #---------------------form the base case gen dictionary finished here-------------------------------
    
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
    isvfiles = ['InitCase'] # making sure InitCase is in the beginning 
    for file in ziplist:
        if '.isv' in file:
			if not file == 'InitCase.isv':
				isvfiles.append(file[:-4])
        if '.sav' in file:
            savefile = file
    #pywin.debugger.set_trace()
    # Get contingency names from excel sheet
    workbook_accc = xlrd.open_workbook(xlsfile)
    sheet_bf = workbook_accc.sheet_by_name('Contingency Events')
    
    cases_delta_dict = {}

    # remove exising solution files
    solutin1file = address + '\\' +case  + '_solution1.txt'
    if os.path.exists(solutin1file):
        os.remove (solutin1file)
    solutin2file = address + '\\' +case + '_solution2.txt'
    if os.path.exists(solutin2file):
        os.remove (solutin2file)    
    
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
            ierr = psspy.getcontingencysavedcase(Zip, isvfile)
        else:
            cont = 'InitCase'
			#ierr = psspy.getcontingencysavedcase(Zip, isvfile)  
        #psspy.fnsl([0,0,0,1,1,0,0,0])
        # extract data for solution 1 and solution 2
        # bus section
        ierr, iarray = psspy.abusint(-1, 1, 'NUMBER')
        vbusno = iarray[0]   # this array has all the bus number
        print "type:", type(vbusno)
        ierr, rarray = psspy. abusreal(-1, 1, 'PU')  
        vbusmag = rarray[0]  # this array has all the bus voltage magnitude
        ierr, rarray = psspy. abusreal(-1, 1, 'ANGLED')  
        vbusangle = rarray[0]      # this array has all the bus voltage angle, in ardians
        
        # generator section
        ierr, iarray = psspy.amachint(-1, 4, 'NUMBER')
        vgenbusno = iarray[0] # this array has all the generator's bus number, including both in-service and out-service
        ierr, iarray = psspy.amachint(-1, 4, 'STATUS')
        vgenstatus = iarray[0] # this array has all the generator's status: in-service (1) and out-service (0)
        ierr, carray = psspy.amachchar(-1, 4, 'ID')
        vgenid = carray[0] # this array has all the generator's ID, string
        ierr, rarray = psspy.amachreal(-1, 4, ['PGEN','PMAX','PMIN'])
        vgenp = map(lambda (a,b):a*b,zip(vgenstatus,rarray[0] ))    # this array has all the generator's Active power output P, MW
        vgenpmax = map(lambda (a,b):a*b,zip(vgenstatus,rarray[1] ))    # this array has all the generator's Pmax, MW
        vgenpmin = map(lambda (a,b):a*b,zip(vgenstatus,rarray[2] ))    # this array has all the generator's Pmin, MW
        ierr, rarray = psspy.amachreal(-1, 4, ['QGEN','QMAX','QMIN'])
        vgenq = map(lambda (a,b):a*b,zip(vgenstatus,rarray[0] )) # this array has all the generator's Reactive power output Q, MVar
        vgenqmax = map(lambda (a,b):a*b,zip(vgenstatus,rarray[1] )) 
        vgenqmin = map(lambda (a,b):a*b,zip(vgenstatus,rarray[2] ))
        # getting the machines bus voltages
        #ierr = psspy.bsys(11, 0, [0.2,999.0], 0, [], len(vgenbusno), vgenbusno,0, [], 0, [])
        #ierr, rarray = psspy. abusreal(11, 2, 'PU')
        #vgenbusvpu = rarray[0]
        #ierr, rarray = psspy. abusint(11, 2, 'NUMBER')
        #vgenbusnounique = rarray[0]
        # switched shunts section
        ierr, iarray = psspy.aswshint(-1, 4, 'NUMBER')
        swshuntbusno = iarray[0] # this array has all the switched shunts bus number
        ierr, iarray = psspy.aswshint(-1, 4, 'STATUS')
        swshuntstatus = iarray[0] # this array has all the switched shunts status
    
        ierr, rarray = psspy.aswshreal(-1, 4, ['BSWNOM','BSWMAX','BSWMIN'])
        swshunt_inival = rarray[0] # this array has all the switched shunts values
        swshunt_max = rarray[1] 
        swshunt_min = rarray[2] 
        
        # form the switched shunt dictionary
        shunt_dict = {}
        vbus_shunt_val = []
        for ibustmp in range(0, len(swshuntbusno)):
            shuntkeytmp = swshuntbusno[ibustmp]
            if 1:
                # impose Bmin and Bmax on extracted values
                if swshunt_inival[ibustmp]>round(100*swshunt_max[ibustmp])/100:
                    swshunt_inival[ibustmp]=round(100*swshunt_max[ibustmp])/100 - eps
                if swshunt_inival[ibustmp]<round(100*swshunt_min[ibustmp])/100:
                    swshunt_inival[ibustmp]=round(100*swshunt_min[ibustmp])/100 + eps
                
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
        totalgendeltamw = 0.0
        totalgendeltacount = 0
        deltatmp_swing = 0.0
        genloss = 0.0
        swing_count = 0.0
        for igentmp in range(0, len(vgenbusno)):
            genbuskeytmp = str(vgenbusno[igentmp])+'-'+vgenid[igentmp].strip()
            genptmp = vgenp[igentmp] * vgenstatus[igentmp]
            
            basegenp = basecase_gen_dict[genbuskeytmp]
            deltatmp = (genptmp - basegenp)/gendroop[genbuskeytmp]
            deltatmpmw = genptmp - basegenp

            # check the amount of generation lost
            basegenstat = basecase_gen_dict_stat[genbuskeytmp]
            if basegenstat == 1 and vgenstatus[igentmp] == 0:
                genloss = genloss + basegenp
            
            # Check the amount of delta that appears in slack bus
            if genbuskeytmp in swing_gen:
                deltatmp_swing = deltatmp_swing + deltatmpmw
                swing_count = swing_count+1.0
                #print('!!!!testout-----------------Swing Gen is ' + genbuskeytmp + ': ' + str(deltatmp))
                
            vgen_delta_dict.update({genbuskeytmp:deltatmp})                    
            tmpstr = 'GEN-'+genbuskeytmp
                            

            #print ('contgen_info')
            #print contgen_info

            if cont in cont_gen_dict.keys():                    
                contgen_info = cont_gen_dict[cont]
                # machine is excluded from the calculation if tripped 
                if  str(vgenbusno[igentmp]) == contgen_info[0] and vgenid[igentmp].strip() == contgen_info[1]:
                    totalgendeltacount = totalgendeltacount - 1
                    deltatmp = 0.0
                    deltatmpmw = 0.0
                    print('!!!!testout----------------- Gen outage at: ' + tmpstr + 'delta =' + str(deltatmp) )
                    
            # if machine is at Pmin or Pmax it should be excluded from delta calculation
            if (vgenp[igentmp] == vgenpmax[igentmp] or vgenp[igentmp] == vgenpmin[igentmp]) and vgenstatus[igentmp]!=0 :
                totalgendeltacount = totalgendeltacount - 1
                deltatmp = 0.0
                #deltatmpmw = 0.0
                print('!!!!testout----------------- Gen is at Pmin or Pmax: ' + tmpstr )

            # machine should also be excluded if out of service
            #basegenstat = basecase_gen_dict_stat[genbuskeytmp]
            if basegenstat == 0 and vgenstatus[igentmp] == 0:
                totalgendeltacount = totalgendeltacount - 1
                
            totalgendelta = totalgendelta + deltatmp
            totalgendeltamw = totalgendeltamw + deltatmpmw
            totalgendeltacount = totalgendeltacount + 1

        totalgendeltamean =  totalgendelta/totalgendeltacount
        
        # check if delta is seen in slack only and if yes re-adjust slack and run inlf
        if (totalgendeltamw>genloss and abs(deltatmp_swing)>1.0) or (deltatmp_swing<0.0):
            #need to modify slack generation back to base case
            for igentmp in range(0, len(vgenbusno)):
                ierr, vgenbustype = psspy.busint(vgenbusno[igentmp] ,'TYPE')
                if vgenstatus[igentmp]==1 and vgenbustype==3:
                    genbuskeytmp = str(vgenbusno[igentmp])+'-'+vgenid[igentmp].strip()
                    slackdroop = gendroop[genbuskeytmp]
                    
                    ierr = psspy.machine_chng_2(vgenbusno[igentmp], vgenid[igentmp], realar1=vgenp[igentmp]-deltatmp_swing/swing_count+0.95*totalgendeltamean*slackdroop)

            ierr = psspy.inlf_2([1,0,0,0,1,0,0,0], fileINL)
            
            # now redo the data collection
            # extract data for solution 1 and solution 2
            # bus section
            ierr, iarray = psspy.abusint(-1, 1, 'NUMBER')
            vbusno = iarray[0]   # this array has all the bus number
            print "type:", type(vbusno)
            ierr, rarray = psspy. abusreal(-1, 1, 'PU')  
            vbusmag = rarray[0]  # this array has all the bus voltage magnitude
            ierr, rarray = psspy. abusreal(-1, 1, 'ANGLED')  
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
            ierr, iarray = psspy.amachint(-1, 4, 'TYPE')
            vgenbustype = iarray[0] # this array has all the generator's bus number, including both in-service and out-service
            # getting the machines bus voltages
            #ierr = psspy.bsys(11, 0, [0.2,999.0], 0, [], len(vgenbusno), vgenbusno,0, [], 0, [])
            #ierr, rarray = psspy.abusreal(11, 2, 'PU')
            #vgenbusvpu = rarray[0]
            #ierr, rarray = psspy. abusint(11, 2, 'NUMBER')
            #vgenbusnounique = rarray[0]
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
                            
                # impose Bmin and Bmax on extracted values
                if 1:
                    if swshunt_inival[ibustmp]>round(100*swshunt_max[ibustmp])/100:
                        swshunt_inival[ibustmp]=round(100*swshunt_max[ibustmp])/100 - eps
                    if swshunt_inival[ibustmp]<round(100*swshunt_min[ibustmp])/100:
                        swshunt_inival[ibustmp]=round(100*swshunt_min[ibustmp])/100 + eps
                    
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
            deltatmp_swing = 0.0
            totalgendeltamw = 0.0
            totalgendeltacount = 0
            genloss = 0.0

            for igentmp in range(0, len(vgenbusno)):

                genbuskeytmp = str(vgenbusno[igentmp])+'-'+vgenid[igentmp].strip()
                genptmp = vgenp[igentmp] * vgenstatus[igentmp]
                
                basegenp = basecase_gen_dict[genbuskeytmp]
                deltatmp = (genptmp - basegenp)/gendroop[genbuskeytmp]
                deltatmpmw = genptmp - basegenp

                # check the amount of generation lost
                basegenstat = basecase_gen_dict_stat[genbuskeytmp]
                if basegenstat == 1 and vgenstatus[igentmp] == 0:
                    genloss = genloss + basegenp
                
                # Check the amount of delta that appears in slack bus
                if genbuskeytmp in swing_gen:
                    deltatmp_swing = deltatmp_swing + deltatmpmw
                    swing_count = swing_count+1.0
                    #print('!!!!testout-----------------Swing Gen is ' + genbuskeytmp + ': ' + str(deltatmp))

                
                vgen_delta_dict.update({genbuskeytmp:deltatmp})                    
                tmpstr = 'GEN-'+genbuskeytmp
                                
                if cont in cont_gen_dict.keys():
                    contgen_info = cont_gen_dict[cont]   
                    # machine is excluded from the calculation if tripped 
                    if str(vgenbusno[igentmp]) == contgen_info[0] and vgenid[igentmp].strip() == contgen_info[1] :
                        totalgendeltacount = totalgendeltacount - 1
                        deltatmp = 0.0
                        deltatmpmw = 0.0
                        print('!!!!testout-----------------Outrage Gen is ' + tmpstr + ': ' + str(deltatmp))

                # if machine is at Pmin or Pmax it should be excluded from delta calculation
                if (vgenp[igentmp] == vgenpmax[igentmp] or vgenp[igentmp] == vgenpmin[igentmp]) and vgenstatus[igentmp]!=0 :
                    totalgendeltacount = totalgendeltacount - 1
                    deltatmp = 0.0
                    #deltatmpmw = 0.0
                    print('!!!!testout----------------- Gen is at Pmin or Pmax: ' + tmpstr )

                # machine should also be excluded if out of service
                #basegenstat = basecase_gen_dict_stat[genbuskeytmp]
                if basegenstat == 0 and vgenstatus[igentmp] == 0:
                    totalgendeltacount = totalgendeltacount - 1

                totalgendelta = totalgendelta + deltatmp
                totalgendeltamw = totalgendeltamw + deltatmpmw
                totalgendeltacount = totalgendeltacount + 1

            totalgendeltamean =  totalgendelta/totalgendeltacount   	

            
        
        #totalgendelta is the delta value for the case    
        cases_delta_dict.update ({cont:totalgendeltamean})  

        # impose Pmin, Pmax, Qmin, Qmax on data extracted
        if 1:
            for igentmp in range(0,len(vgenid)):
                vgenqmax[igentmp] = round(10000*vgenqmax[igentmp])/10000
                vgenqmin[igentmp] = round(10000*vgenqmin[igentmp])/10000
                vgenpmax[igentmp] = round(10000*vgenpmax[igentmp])/10000
                vgenpmin[igentmp] = round(10000*vgenpmin[igentmp])/10000
                
                if vgenq[igentmp] > vgenqmax[igentmp]:
                    vgenq[igentmp] = vgenqmax[igentmp]
                if vgenq[igentmp] < vgenqmin[igentmp]:
                    vgenq[igentmp] = vgenqmin[igentmp]
                if vgenp[igentmp] > vgenpmax[igentmp]:
                    vgenp[igentmp] = vgenpmax[igentmp]
                if vgenp[igentmp] < vgenpmin[igentmp]:
                    vgenp[igentmp] = vgenpmin[igentmp]


        if 0:
            for ibustmp in range(0,len(vbusno)):
                vbusmag[ibustmp] = round(1000000*vbusmag[ibustmp])/1000000
        
        if cont!='InitCase':
            for k in range(0,len(vgenbusno)):
                ierr, genvpu = psspy.busdat(vgenbusno[k],'PU')
                basegenvpu = basebusvpu_dict[vgenbusno[k]]
                if genvpu<basegenvpu:
                    if (basegenvpu-genvpu)>(vgenqmax[k]-vgenq[k])/100:
                        vgenq[k]=vgenqmax[k]
                        #vgenbusvpu[k] is same
                    else:
                        vbusindx = vbusno.index(vgenbusno[k])
                        vbusmag[vbusindx] = basegenvpu
                        #vgenbusvpu[k] = vbasecasegenbusvpu[k]
                        # vgenq[k] is same
                if genvpu>basegenvpu:
                    if (genvpu-basegenvpu)>(vgenq[k]-vgenqmin[k])/100:
                        vgenq[k] = vgenqmin[k]
                        # vgenbusvpu[k] is same
                    else:
                        vbusindx = vbusno.index(vgenbusno[k])
                        vbusmag[vbusindx] = basegenvpu
                        #vgenbusvpu[k] = vbasecasegenbusvpu[k]
                        #vgenq[k] is same

            # updating the bus voltages based on the generators bus voltages
            #for k in range(0,len(vgenbusnounique)):
            #    vbusindx = vbusno.index(vgenbusnounique[k])
            #    #print vbustmp
            #    vbusmag[vbusindx] = vgenbusvpu[k]

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
        '''
        if os.path.exists(solutin1file):
            os.remove (solutin1file)
        if os.path.exists(solutin2file):
            os.remove (solutin2file)
        '''

        System = COMPET_FORM(cont,solutin1file, solutin2file, vbusno, vbusmag,vbusangle,vgenbusno,vgenid, vgenp, vgenq, vbus_shunt_val,totalgendeltamean)
        
        # save each post-contingency case in both raw and sav format
        # psspy.save(scopfaddress + '\\' + caseX +'\\'+caseX +"_" + str(cont)+'.sav')
        psspy.rawd_2(0,1,[1,1,1,0,0,0,0],0,scopfaddress + '\\' + case + '\\' + case +"_" + str(cont)+'.raw')
    
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

    if 0: # no need for this since we are saving a separate copy in each subfolder
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

rawfile = 'case.raw'
confile = 'case.con'
inlfile = 'case.inl'
monfile = 'All.mon'
subfile = 'All_SDET.sub'
GOValid_func(rawfile,confile,inlfile,monfile,subfile)

GOValid_func('case14.raw','case14.con','case14.inl','All.mon','All_SDET.sub')
'''
