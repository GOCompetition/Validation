#!/bin/sh

# can somebody write a WINDOWS batch file version of this?
# the validation script only runs on windows,
# so this script is just illustrative

# users of the validation script should feel free to write their own
# versions of this script or modify it as needed.
# just please do not upload changes to this script to the repo

case_dir='./case14small/'
raw=$case_dir'case14.raw'
#rop=$case_dir'case14.rop'
con=$case_dir'case14.con'
inl=$case_dir'case14.inl'
#sol1=$case_dir'sol1.txt'
#sol2=$case_dir'sol2.txt'
mon=$case_dir'All.mon'
sub=$case_dir'All_SDET.SUB'

# run it
#python ./validation.py "$raw" "$rop" "$con" "$inl" "$sol1" "$sol2"
python ./validation.py "$raw" "$con" "$inl" "$mon" "$sub"
