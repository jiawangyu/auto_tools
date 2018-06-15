#coding=utf-8
# replace.py

import os
import win32com.client
import pywintypes
import pythoncom
import win32api

win32com.client.gencache.EnsureDispatch ("Word.Application")
#---------------- common --------------------------------------
def init(visible=False):
    word = win32com.client.DispatchEx('Word.Application')
    word.Visible = visible
    word.DisplayAlerts = False
    return word

def open(word, filename):
    return word.Documents.Open(filename)

def replaceAll(word, oldStr,newStr):
    find = word.Selection.Find
    find.ClearFormatting()
    find.Replacement.ClearFormatting()
    print("%s => %s" %(oldStr, newStr))
    find.Execute(oldStr, False, False, False, False, False, True, 1, True, newStr, 2)   

def saveAs(word, filename, delete_existing=True):
    if delete_existing and os.path.exists(filename):
        os.remove(filename)
    word.ActiveDocument.SaveAs(FileName=filename)

def quit(word):
    word.ActiveDocument.Close()
    return word.Quit()

#-------------------- test -----------------------------
#1. insert text
def insert_test(doc, str):
    myRange = doc.Range(0,0)
    myRange.InsertBefore(str)

#2. text replace
def replace_test(word, oldstr, newstr):
    replaceAll(word, oldstr, newstr)

#3. style set
def style_test(doc):
    #myRange=doc.Range(doc.Content.Start, doc.Content.End)
    myRange=doc.Range(0, 0)
    myRange.InsertAfter('wdStyleHeading1 test!')
    myRange.Style = win32com.client.constants.wdStyleHeading1
    myRange.InsertAfter('wdStyleHeading2 test!')
    myRange.Style = win32com.client.constants.wdStyleHeading2
    myRange.InsertAfter('wdStyleHeading3 test!')
    myRange.Style = win32com.client.constants.wdStyleHeading3

#4. header/footer
def headers_test(word, oldstr, newstr):
    word.ActiveWindow.ActivePane.View.SeekView = 9 #9: 页眉； 10: 页脚
    word.Selection.ParagraphFormat.Alignment = 0
    word.Selection.Text = 'June@SMEE'
    word.ActiveWindow.ActivePane.View.SeekView = 0 # 释放焦点，返回主文档

#5. header/footer image

#6. heading/title
def heading_test(word, str):
    for section in word.ActiveDocument.Sections:
        for header in section.Headers:
            print(header.Range.Text)
            header.Range.Text="aaaaaaaaaaaaaaaaaaabbbbbbbbbbbbbbbb"
    header_range= doc.Sections(1).Headers(win32.constants.wdHeaderFooterPrimary).Range
    header_range.ParagraphFormat.Alignment = win32.constants.wdAlignParagraphCenter
    header_range.Font.Bold = True
    header_range.Font.Size = 12
    header_range.Text = "cccccccccccccc"
#7. watermark

#8. page set
def page_test(doc):
    doc.PageSetup.Orientation = 1 # Make some Setup to the Document:
    doc.PageSetup.LeftMargin = 20
    doc.PageSetup.TopMargin = 20
    doc.PageSetup.BottomMargin = 20
    doc.PageSetup.RightMargin = 20
    doc.Content.Font.Size = 11
    doc.Content.Paragraphs.TabStops.Add (100)
    doc.Content.Text = "Hello, I am a text!"
    doc.Content.MoveEnd

#9 bookmarks
def bookmarks_test(doc, str):
    try:
        myRange=doc.Range(doc.Content.Start, doc.Content.End)
        myRange.Bookmarks.Add1(Name=str)
    except pythoncom.com_error as error:
        print(error.strerror)

#10 table
def table_test(doc):
    total_column = 3
    total_row = len(compound_name)+1
    rng = doc.Range(0,0)
    rng.ParagraphFormat.Alignment = win32com.client.constants.wdAlignParagraphCenter
    table = doc.Tables.Add(rng,total_row, total_column)
    table.Borders.Enable = True
    if total_column > 1:
        table.Columns.DistributeWidth()
    #table title
    table.Cell(1,1).Range.InsertAfter("title1")
    table.Cell(1,2).Range.InsertAfter("title2")
    table.Cell(1,3).Range.InsertAfter("title3")

#11 image
    frame_max_width= 167 # the maximum width of a picture
    frame_max_height= 125 # the maximum height of a picture
    #
    for index, filename in enumerate(filenames): # loop through all the files and folders for adding pictures
        if os.path.isfile(os.path.join(os.path.abspath("."), filename)): # check whether the current object is a file or not
            if filename[len(filename)-3: len(filename)].upper() == 'PNG': # check whether the current object is a JPG file
                cell_column= index % total_column + 1
                cell_row = index / total_column + 2

                cell_range= table.Cell(cell_row, cell_column).Range
                cell_range.ParagraphFormat.LineSpacingRule = win32.constants.wdLineSpaceSingle
                cell_range.ParagraphFormat.SpaceBefore = 0 
                cell_range.ParagraphFormat.SpaceAfter = 3

                #this is where we are going to insert the images
                current_pic = cell_range.InlineShapes.AddPicture(os.path.join(os.path.abspath("."), filename))

# 设置标题样式
def heading_test(doc):
    doc.Styles(win32com.client.constants.wdStyleHeading1).Font.Name = "Arial"
    doc.Styles(win32com.client.constants.wdStyleHeading1).Font.Bold = False
    doc.UpdateStyles

#-------------------- main -----------------------------
def main():
    word = init()
    filename="demo.docx"

    doc = open(word, os.path.abspath(filename))
    try:
        #style_test(doc)
        #insert_test(doc, 'This is June!')
        #replace_test(word, 'aa', 'bb')
        #headers_test(word, 'aa', 'bb')
        ##table_test(doc)
        #heading_test(word, 'heading test')
        ##bookmarks_test(doc, 'bookmarks test')
        #page_test(doc)
        heading_test(doc)

        saveAs(word, os.path.abspath("new_"+filename))
    finally:
        quit(word)

if __name__ == '__main__':
    main()