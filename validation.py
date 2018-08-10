'''
syntax
python validation.py <raw> <con> <inl> <mon> <sub>

'''

import argparse

def run_main(raw, con, inl, mon, sub):
    """run with given file names"""

    # Ahmad please put (or import/call) your code here
    
def run():
    """parse file names and run"""

    parser = argparse.ArgumentParser(description='Run the GOComp PSSE based validation tool on a problem instance')
    
    parser.add_argument('raw', help='raw - complete path and file name to a RAW file')
    parser.add_argument('con', help='con - complete path and file name to a CON file')
    parser.add_argument('inl', help='inl - complete path and file name to a INL file')
    parser.add_argument('mon', help='mon - complete path and file name to a MON file')
    parser.add_argument('sub', help='sub - complete path and file name to a SUB file')
    
    args = parser.parse_args()
    
    try:
        raw = args.raw
        con = args.con
        inl = args.inl
        mon = args.mon
        sub = args.sub
    except:
        print "exception in parsing the validation command"
        raise
    else:
        run_main(raw, con, inl, mon, sub)

if __name__ == '__main__':
    run()
