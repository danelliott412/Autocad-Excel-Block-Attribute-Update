# This program was written by Dan Elliott to import blocks with attributes calculated in Excel into AutoCAD.
# The first Block entered needs to be set up or pulled from a block specific .dwg file then it can referenced in your
# current AutoCAD file
from pyautocad import Autocad, APoint       # acad2
# from openpyxl import load_workbook
import win32com.client                      # acad1
import array
from fractions import Fraction
import re
import time
import xlwings as xw


app = xw.apps.active  # set to read active Excel workbook as this code is called from workbook macro
ws = xw.sheets.active

acad1 = win32com.client.Dispatch("AutoCAD.Application")   # setup for win32com
acad2 = Autocad(create_if_not_exists=True)                # setup for pyautocad

doc1 = acad1.ActiveDocument  # Document object
doc2 = acad2.ActiveDocument
# First instance must be created or imported from block drawing
BLOCK = r'C:string\to\block\BLOCKNAME.dwg'

AA = array.array('f', [])   # set up arrays with floats
A = array.array('f', [])
B = array.array('f', [])
C = array.array('f', [])
D = array.array('f', [])

cellcount = 0
rn = 15
for rn in range(15, 66):
    rn = str(rn)
    if ws['V' + rn].value is None:
        break
    else:
        cellcount += 1
        AA.append(ws['V' + rn].value)
        A.append(ws.range('W' + rn).value)
        B.append(ws.range('X' + rn).value)
        C.append(ws.range('Y' + rn).value)
        D.append(ws.range('Z' + rn).value)


startpnt = doc1.Utility.GetPoint()
fistinsert = APoint(startpnt)

# pull from file to initialize first block reference
block_ref1 = doc2.PaperSpace.InsertBlock(fistinsert, BLOCK, 1, 1, 1, 0)
time.sleep(0.15)

xpnt = 1.5
ypnt = 0
rng = range(cellcount - 1)

for i in rng:
    # translate insertion point
    transpnt = APoint(startpnt) + APoint(xpnt, ypnt)
    # Insert block (insert point, "BLOCKNAME", scale)
    block_ref2 = acad2.ActiveDocument.PaperSpace.InsertBlock(transpnt, "BLOCKNAME", 1, 1, 1, 0)
    xpnt += 1.5
    if xpnt == 6:
        ypnt -= 1.75
        xpnt = 0
        
blockcount = 0   # counter for blocks

# iterate through all objects (entities) in the currently opened drawing
for entity in doc1.PaperSpace:
    name = entity.EntityName
    if name == 'AcDbBlockReference':
        HasAttributes = entity.HasAttributes
        # Checks for entities in AutoCAD with specific block name to edit attributes
        if HasAttributes and entity.Name == 'BLOCKNAME':
            entity.Layer = 'LAYER NAME'
            # print(entity.ObjectID)
            for attrib in entity.GetAttributes():

                Aval = A[blockcount]
                Bval = B[blockcount]
                Cval = C[blockcount]
                Dval = D[blockcount]

                # update attributes
                
                if attrib.TagString == 'AA':
                    attrib.TextString = str(AA[0])
                if attrib.TagString == 'A':
                    attrib.TextString = str(int(Aval))
                if attrib.TagString == 'B':
                    # converts floats below from decimal to show as fraction with inch symbol
                    attrib.TextString = re.sub(' 0', '', str(int(Bval)) + " " + str(Fraction(Bval % 1)) + '"')
                if attrib.TagString == 'C':
                    attrib.TextString = re.sub(' 0', '', str(int(Cval)) + " " + str(Fraction(Cval % 1)) + '"')
                if attrib.TagString == 'D':
                    attrib.TextString = re.sub(' 0', '', str(int(Dval)) + " " + str(Fraction(Dval % 1)) + '"')
                    attrib.Update()
                    blockcount += 1
