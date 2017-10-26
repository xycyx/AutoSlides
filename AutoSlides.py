#!/usr/bin/env python2
# -*- coding: utf-8 -*-
"""
Created on Tue Oct 17 00:56:37 2017

@author: yxcheng
"""
import sys
import argparse

#print 'Number of arguments:', len(sys.argv), 'arguments.'
#print 'Argument List:', str(sys.argv)


import sys, getopt
from pptx import Presentation
#import re
#import os
from utils import *

def main(argv):

    inputfile = ''
    outputfile = ''
    try:
        opts, args = getopt.getopt(argv,"hi:o:",["ifile=","ofile="])
    except getopt.GetoptError:
        print ('python AutoSlides.py -i <inputfile> -o <outputfile>')
        sys.exit(2)
    for opt, arg in opts:
        if opt == '-h':
            print ('python AutoSlides.py -i <inputfile> -o <outputfile>')
            sys.exit()
        elif opt in ("-i", "--ifile"):
            inputfile = arg
        elif opt in ("-o", "--ofile"):
            outputfile = arg
    print ('Input file is "', inputfile)
    print ('Output file is "', outputfile)

#####################
#%%
    prs = Presentation()

#   inputfile = './example2/UWANG5000-977/'
    file_lists = search_images_in_folder(inputfile)
    df = analysis_filelists(file_lists)

    prs = add_a_table(prs, 12, 2, 'Summary', df)

    #scan_time = '20161123161757'
    #prs = build_the_slide(prs, df, 'Retina', scan_time)
    prs = create_slides(df, prs)

#   outputfile = 'test3.pptx'
    prs.save(outputfile)
    print('Slides saved!')

#%%
if __name__ == "__main__":
   main(sys.argv[1:])
#if __name__ == "__main__":
#    args = parse_args()
#
#    print(args)
