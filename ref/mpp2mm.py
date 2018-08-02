#!/usr/bin/env python
#

"""Create a FreeMind (.mm) mind map from a/an MS Project XML file.

usage: mpp2mm -i <mppfile> -o <mmfile> -a -c
Options and arguments:
-a         : generate (some) Attributes
-c         : generate icons for completed tasks
-p         : generate arrows linking predecessors
-i mppfile : input MS Project mpp or xml file
-o mmfile  : output FreeMind map (name defaults to input filename)

Copyright (C) 2009, AnGus King

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 2 of the License, or
any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <http://www.gnu.org/licenses/>.

"""

__author__ = "AnGus King"
__copyright__ = "Copyright (C) 2009, AnGus King" 
__license__ = "GPLv2 (http://www.gnu.org/licenses/)" 
__version__ = "1.0" 

import getopt
from mppxml2mm import Mppxml2Mm
import sys
import win32com.client
import xml.etree.ElementTree as ET


class Mpp2Mm:
    
    def __init__(self):
        self.mmn = [] # create an empty node tree
        self.mmn.append("")
        self.lst_lvl = 0 # last level is 0
        self.do_attr = False # by default don't create Attributes
        self.do_flag = False # by default don't flag completion 
        self.do_prior = False # by default don't produce predecessor arrows       

    def open(self, infile):
        """ Open the .mpp or .xml file and create a tree """
        if infile.endswith('.xml'):
            self.proj = None
            try:
                self.tree = ET.parse(infile)
                return self.tree
            except Exception, e:
                print "Error opening file",e
                usage()
        else:
            try:
                self.mpp = win32com.client.Dispatch("MSProject.Application")
                self.mpp.Visible = False
                self.mpp.FileOpen(infile)
                self.proj = self.mpp.ActiveProject
                return self.proj
            except Exception, e:
                print "Error opening file",e
                usage()

    def write(self, outfile, attr, flag, prior):
        """ Create the .mm file :-) """

        def mps(elem):
            return "{http://schemas.microsoft.com/project}"+elem

        def dys(msdurn):
            durn = ""
            if msdurn is not None:
                if len(msdurn) > 3:
                    if msdurn[0:2] == "PT":
                      hrs = msdurn[2:].split("H") # forget about minutes , ...
                      durn = float(hrs[0])/8 
                      return str(durn)[0:-2] + "d"
            return durn

        self.do_attr = attr
#       create the mindmap as version 8.1, or 9.0 if generating attribute tags
        if self.do_attr:
            self.mm = ET.Element("map",version="0.9.0")
        else:
            self.mm = ET.Element("map",version="0.8.1")
        self.do_flag = flag
        self.do_prior = prior
        lvli = 0 # xml only
        last_lvl = 0
        last_durn = "" # xml only
        if self.proj is not None: # .mpp file
            self.mmn[0] = ET.SubElement(self.mm, "node", TEXT=self.proj.Project)
            if self.do_attr:
                txt = self.proj.Comments
                if txt is not None:
                    if txt != "":
                        txt = txt.encode("ascii","replace") # in case of any nasties
                        nte = ET.fromstring("<richcontent TYPE=\"NOTE\"><html><head> </head><body><p>" \
                            + txt.strip() + "</p></body></html></richcontent>")
                        self.mmn[lvli].append(nte)
                txt = str(self.proj.ProjectStart)
                mth = txt.partition("/") # get month
                day = mth[2].partition("/")
                yer = day[2].partition(" ")
                txt = "20" + yer[0] + "-" + mth[0] + "-" + day[0]
                self.mmn[0].append(ET.Element("attribute", \
                    NAME="prj-Author", VALUE=self.proj.Author))
                self.mmn[0].append(ET.Element("attribute", NAME="prj-StartDate", VALUE=txt))
            for task in self.proj.Tasks:
                if task is not None:
                    lvli = task.OutlineLevel
                    txt = task.Name 
                    if lvli > self.lst_lvl:
                        self.mmn.append("")
                        self.lst_lvl = lvli
                    self.mmn[lvli] = ET.SubElement(self.mmn[lvli-1], "node",TEXT=txt, \
                        ID="ID_" + str(task.ID))
                    if self.do_attr:
                        if task.Notes != "":
                            nte = ET.fromstring("<richcontent TYPE=\"NOTE\"><html><head> </head><body><p>" \
                                + task.Notes.strip() + "</p></body></html></richcontent>")
                            self.mmn[lvli].append(nte)
                        if not task.Estimated:
                            self.mmn[lvli].append(ET.Element("attribute", NAME="tsk-Estimated",\
                                VALUE="0"))
                        if self.do_flag:
                            self.mmn[lvli].append(ET.Element("attribute", NAME="tsk-PercentComplete",\
                                VALUE=str(task.PercentComplete)))
                        if task.OutlineChildren.Count == 0: # duration for lowest level only
                            txt = str(task.Duration / 480) + "d" # convert to days
                            self.mmn[lvli].append(ET.Element("attribute", NAME="tsk-Duration", \
                                VALUE=txt))
                    if self.do_flag:
                        txt = task.PercentComplete
                        if txt == 100:
                            self.mmn[lvli].append(ET.Element("icon", BUILTIN="button_ok"))
                    if self.do_prior: # predecessors
                        txt = task.Predecessors # task.Successors
                        while len(txt) > 0:
                            bits = txt.partition(",") # get next predecessor
                            txt = bits[2]
                            try:
                                for i in range(0,self.lst_lvl): # find the parent
                                    for j in range(0, len(self.mmn[i])):
                                        if self.mmn[i][j].get("ID") == "ID_" + str(bits[0]):
                                            self.mmn[i][j].append(ET.Element("arrowlink", COLOR="#0000ff", \
                                                DESTINATION="ID_" + str(task.ID), ENDARROW="Default", \
                                                STARTARROW="None"))
                                            raise StopIteration()
                            except StopIteration:
                                pass
            self.mpp.Quit()
            
        tree = ET.ElementTree(self.mm)
        tree.write(outfile)
        return

def usage():
    print __doc__
    sys.exit(-1)

def main():
    try:
        opts , args = getopt.getopt(sys.argv[1:], "i:o:hacp", \
            ["help", "input=", "output="])
    except getopt.GetoptError, err:
        # print help information and exit:
        print str(err) # will print something like "option -x not recognized"
        usage()
    input = None
    output = None
    attr = False
    complete = False
    prior = False
    for o, a in opts:
        if o == "-a": # wants attributes printed
            attr = True
        elif o == "-c": # wants completion flags
            complete = True
        elif o == "-p": # wants predecessor arrows
            prior = True     
        elif o in ("-h", "--help"):
            usage()
        elif o in ("-i", "--input"):
            input = a
        elif o in ("-o", "--output"):
            output = a
        else:
            assert False, "unhandled option"
    if input == None:
        print "Input file required"
        usage() 
    if not (input.endswith('.xml') or input.endswith('.mpp')):
        print "Input file must end with '.mpp' or '.xml'"
        usage()
    if output == None:
        output = input[:-3] + "mm"
    if not output.endswith('.mm'):
        print "Output file must end with '.mm'"
        usage()
    if input.endswith('.xml'):
        mpp2mm = Mppxml2Mm()
    else:
        mpp2mm = Mpp2Mm()
    tree = mpp2mm.open(input)
    mpp2mm.write(output,attr,complete,prior)
    print output + " created."


if __name__ == "__main__":
    main()
