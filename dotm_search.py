#!/usr/bin/env python
"""
Given a directory path, this searches all files in the path for a given text string 
within the 'word/document.xml' section of a MSWord .dotm file.
"""
import sys
import os
import glob
import zipfile
import re
import argparse


# Your awesome code begins here!

def chosenDir(text, folder):
    os.chdir(folder)
    directory = os.getcwd()
    files = glob.glob('*.dotm')
    for f in files:
        zf = zipfile.ZipFile(f, "r")
        data = zf.read("word/document.xml")
        for index, i in enumerate(data):
            if i == text:
                print """
                Match found in file """, directory + "/" + f, """

                ...""", data[index - 40: index] + data[index:index + 40]
    os.chdir('../')


def currentDir(text):
    directory = os.getcwd()
    files = glob.glob('*.dotm')
    for f in files:
        zf = zipfile.ZipFile(f, "r")
        data = zf.read("word/document.xml")
        for index, i in enumerate(data):
            if i == text:
                print """
                Match found in file """, directory + "/" + f, """

                ...""", data[index - 40: index] + data[index:index + 40]
   
def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('text', help='text to search within each dotm file')
    parser.add_argument('--dir', help='directory path containing dotm files to search. Default is cwd.')
    args = parser.parse_args()

    text = args.text
    folder = args.dir
    if folder:
        chosenDir(text, folder)
    elif not folder:
        currentDir(text)
    else:
        print 'unknown option: ' 
        sys.exit(1)

if __name__ == '__main__':
  main()


