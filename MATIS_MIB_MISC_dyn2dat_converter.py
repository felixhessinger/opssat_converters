'''
Author: Felix Hessinger
Email: felix.hessinger@gmail.com
Phone: +49-1578-7373424
Date: 10/04/2019

This script converts MISCcontext.dyn files from SCOS to MISCconfig.dat files. MISCconfig.dat files are used by MATIS
to read parameters from SCOS in MATIS and to react accordingly.
(The script can also be used for other .dyn to .dat conversions)

DISCLAIMER: It is just a tool to make life easier. Please check each generated file for errors before implementing it.
            This code might contain overseen bugs.
'''

import argparse
########################
# Start Script
ap = argparse.ArgumentParser()
ap.add_argument("-i", "--input_file_name", required=True, help="path to input file (<name>.dyn)")
ap.add_argument("-o", "--output_file_name", required=False, help="path to output file (<name>.dat)")
args = vars(ap.parse_args())

fcomment = False
fname = False
first_line = True
descriptionMISC = ""
fi = open(args["input_file_name"], "r")
fo = open(args["output_file_name"], "w+")
for line in fi:
    if not first_line:
        if line.startswith("###"):
            fcomment = True
        if not line.startswith("#"):
            fcomment = False
            nameMISC = (str(line.split("\t", 1)[0]) + "\t")
            if "====" in descriptionMISC:
                descriptionMISC = ""
            fo.write(nameMISC + descriptionMISC + "\n")
            descriptionMISC = ""
        if fcomment:
            descriptionMISC += (line.replace("# ", "").replace("#", "").replace("\n", " "))
    first_line = False
fi.close()
fo.close()
# End script
########################
