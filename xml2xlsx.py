# xml2xlsx tool that converts xml to Excel data spreadsheet.
# By Bingyin Hu, 2018/12/02
import xlsxwriter
from lxml import etree
import os

class xml2xlsx(object):
    def __init__(self, xmlDir):
        self.failedSheets = []# a list for sheets not created successfully
        self.tree = etree.parse(xmlDir)
        self.xlsxName = os.path.splitext(os.path.split(xmlDir)[-1])[0] + '.xlsx'
        self.wb = xlsxwriter.Workbook(self.xlsxName)
        self.lvOneEles = self.tree.findall('./') # level 1 elements in the xml tree
        # set formats here
        self.title_format = self.wb.add_format({'bold': True, 'font_size': 16})
        self.subtitle_format = self.wb.add_format({'bold': True, 'font_size': 12})
        self.header_format = self.wb.add_format({'bold': True, 'font_size': 12,
                                            'bg_color': '#dbeff9'})
        self.default = self.wb.add_format({'font_size': 11})

    # run function that kickoffs the job
    def run(self):
        for ele in self.lvOneEles:
            try:
                self.createWS((ele), ele.tag)
                print "Worksheet '%s' created." %(ele.tag)
            except:
                self.failedSheets.append(ele)
                continue
        self.save()

    # nm_run function that kickoffs the job customized for nanomine
    def nm_run(self):
        info = [] # Sample Info sheet for ID, Control_ID, DATA_SOURCE
        count = 0
        for ele in self.lvOneEles:
            # 'ID', 'Control_ID', and 'DATA_SOURCE' go into 'Sample Info' sheet
            if ele.tag == 'ID':
                info.append(ele)
            elif ele.tag == 'Control_ID':
                info.append(ele)
            elif ele.tag == 'DATA_SOURCE':
                info.append(ele)
            elif ele.tag == 'PROPERTIES':
                props = ele.findall('./')
                for ele_prop in props:
                    try:
                        self.createWS((ele_prop), ele_prop.tag + ' Properties')
                        print "Worksheet '%s' created." %(ele_prop.tag + ' Properties')
                    except:
                        self.failedSheets.append(ele_prop)
                        continue
            else:
                try:
                    self.createWS((ele), ele.tag)
                    print "Worksheet '%s' created." %(ele.tag)
                except:
                    self.failedSheets.append(ele)
                    continue
            if len(info) > 0: # now deal with Sample Info
                # if it's the last element in self.lvOneEles or the next element
                # is not one of the three elements that go into 'Sample Info'
                if (count >= len(self.lvOneEles) - 1) or (self.lvOneEles[count + 1].tag not in ['ID', 'Control_ID', 'DATA_SOURCE']): 
                    try:
                        self.createWS(tuple(info), "Sample Info")
                        print "Worksheet 'Sample Info' created."
                        info = []
                    except:
                        self.failedSheets += info
                        info = []
                        continue
            count += 1 # add count
        # end of loop

    # save function that closes the workbook and print failed cases if needed
    def save(self):
        if len(self.failedSheets) > 0:
            print "Worksheets cannot be created for:"
            for ele in self.failedSheets:
                print '\t' + str(ele)
        self.wb.close()

    # createWS function that takes the level-one element and creates worksheet
    def createWS(self, eleTup, sheetname):
        ws = self.wb.add_worksheet(sheetname)
        row = 0
        # write title
        ws.write(row, 0, sheetname, self.title_format)
