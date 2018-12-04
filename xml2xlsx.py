# xml2xlsx tool that converts xml to Excel data spreadsheet.
# By Bingyin Hu, 2018/12/02
import xlsxwriter
from lxml import etree
import os
import collections

class xml2xlsx(object):
    def __init__(self, xmlDir,
                 column = ['description', 'value', 'unit',
                           'uncertainty type', 'uncertainty value', 'data'],
                 ignore = ['ChooseParameter', 'AxisLabel',
                           'xName', 'xUnit', 'yName', 'yUnit'],
                 dataTag = ['data', 'LoadingProfile']):
        self.failedSheets = []# a list for sheets not created successfully
        self.column = column # defines the headers starting from column B, tags by default occupies column A
        self.scalarUncertainty = ['description', 'value', 'unit', 'uncertainty', 'data'] # for NanoMine ScalarUncertainty type
        self.ignore = ignore # defines tags that should be ignored when writing to the worksheet
        self.appendix = collections.OrderedDict() # an ordered dict for saving appended data info
        self.dataTag = dataTag # tags that contains data info
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
        self.def_val = self.wb.add_format({'font_size': 11, 'bg_color': '#dae6b6'})
    # run function that kickoffs the job
    def run(self):
        for ele in self.lvOneEles:
            try:
                self.createWS((ele), ele.tag)
                print "Worksheet '%s' created." %(ele.tag)
            except:
                self.failedSheets.append(ele)
                continue
        self.generateAppendix()
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
        self.generateAppendix()

    # save function that closes the workbook and print failed cases if needed
    def save(self):
        if len(self.failedSheets) > 0:
            print "Worksheets cannot be created for:"
            for ele in self.failedSheets:
                print '\t' + str(ele)
        self.wb.close()

    # createWS function that takes the level-one element and creates worksheet
    def createWS(self, eleTup, sheetname):
        # create worksheet
        ws = self.wb.add_worksheet(sheetname)
        row = 0
        # write title
        ws.write(row, 0, sheetname, self.title_format)
        row += 2 # leave a blank row below the title
        # Sample Info, flatten everything
        if sheetname == 'Sample Info':
            for ele in eleTup:
                for node in ele.iter():
                    if node.text is not None:
                        ws.write(row, 0, node.tag, self.default) # first column
                        ws.write(row, self.getCol('description'),
                                 node.text, self.def_val) # second column
                        row += 1
        # otherwise, common cases
        # if a node has no text, but one of its child node does, write header
        # next to it. 
        else:
            for ele in eleTup:
                # a log shows whether the previous node has text
                prevText = False
                skipDownCount = 0
                # TODO
                for node in ele.iter():
                    # as long as skipDownCount is not 0, skip
                    if skipDownCount > 0:
                        skipDownCount -= 1
                        continue
                    # skip tags specified in the ignore
                    if node.tag in self.ignore:
                        continue
                    # data and LoadingProfile goes into appendix, set skipDownCount
                    if node.tag in self.dataTag:
                        # update appendix and get the appendix sheet name
                        appendixSheetName = self.updateAppendix(node)
                        ws.write(row, self.getCol('data'),
                                 'See worksheet '+ appendixSheetName,
                                 self.def_val)
                        # set skipDownCount
                        skipDownCount = len(node.findall('.//'))
                        row += 1
                        continue
                    # should this be a header line?
                    if (self.childText(node) or self.childSU(node)) and node.tag not in self.scalarUncertainty:
                        ws.write(row, 0, node.tag, self.subtitle_format)
                        # write the header
                        for col in xrange(len(self.column)):
                            ws.write(row, col + 1,
                                     self.column[col], self.header_format)
                        # move to next row
                        row += 1
                        continue
                    # is this a ScalarUncertainty type node?
                    if self.typeSU(node):
                        # write first column
                        ws.write(row, 0, node.tag, self.default)
                        # write the child nodes
                        for child in node.findall('./'):
                            # description
                            if child.tag == 'description':
                                ws.write(row, self.getCol(child.tag),
                                         child.text, self.def_val)
                            # value
                            elif child.tag == 'value':
                                ws.write(row, self.getCol(child.tag),
                                         child.text, self.def_val)                                
                            # unit
                            elif child.tag == 'unit':
                                ws.write(row, self.getCol(child.tag),
                                         child.text, self.def_val)
                            # uncertainty
                            elif child.tag == 'uncertainty':
                                for grandchild in child.findall('./'):
                                    if grandchild.text is not None:
                                        tag = child.tag + ' ' + grandchild.tag
                                        ws.write(row, self.getCol(tag),
                                                 grandchild.text, self.def_val)
                            # data goes into appendix, set skipDownCount
                            if child.tag in self.dataTag:
                                # update appendix and get the appendix sheet name
                                # no need to set skipDownCount for data here,
                                # we set it for the ScalarUncertainty type node
                                # as a whole
                                appendixSheetName = self.updateAppendix(child)
                                ws.write(row, self.getCol('data'),
                                         'See worksheet '+ appendixSheetName,
                                         self.def_val)
                        # set skipDownCount
                        skipDownCount = len(node.findall('.//'))
                        row += 1
                        continue
                    # else, general case
                    if node.text is not None:
                        ws.write(row, 0, node.tag, self.default)
                        ws.write(row, self.getCol('description'),
                                 node.text, self.def_val)
                    else:
                        ws.write(row, 0, node.tag, self.subtitle_format)
                    row +=1
        # set column width at the end
        ws.set_column(0, len(self.column) + 1, 17)

    # update appendix and returns the sheet name of the added data
    # self.Appendix has keys like 'Appendix_1', 'Appendix_2', etc.
    def updateAppendix(self, node):
        info = collections.OrderedDict()
        # if appendix is empty, generate 'Appendix_1'
        if len(self.appendix) == 0:
            sheetname = 'Appendix_1'
        else:
            index = int(self.appendix.keys()[-1].split('_')[-1]) + 1
            sheetname = 'Appendix_%i' %(index)
        # parse the data element
        childs = node.findall('./')
        for child in childs:
            if child.tag == 'description':
                info[child.tag] = child.text
            elif child.tag == 'AxisLabel':
                for grandchild in child.findall('./'):
                    info[grandchild.tag] = grandchild.text
            elif child.tag == 'data':
                XY = []
                for ele in child.iter():
                    if ele.text is not None:
                        XY.append(ele.text)
                # even position to X, odd position to Y
                X = []
                Y = []
                for i in xrange(len(XY)):
                    if i%2 == 0:
                        X.append(XY[i])
                    else:
                        Y.append(XY[i])
                info['X'] = X
                info['Y'] = Y
        # if info is not empty, save it to appendix with sheetname as the key
        if len(info) > 0:
            self.appendix[sheetname] = info
            return sheetname
        # otherwise, return nothing
        return ''

    # generates the data sheets saved in the appendix 
    def generateAppendix(self):
        for data in self.appendix:
            try:
                ws = self.wb.add_worksheet(data)
                row = 0
                # write title
                ws.write(row, 0, data, self.title_format)
                row += 2 # leave a blank row below the title
                # expand the appendix dict and write data
                appDict = self.appendix[data]
                if 'description' in appDict:
                    ws.write(row, 0, 'description', self.subtitle_format)
                    ws.write(row, 1, appDict['description'], self.def_val)
                    row += 2
                # write a column of 'X'
                if 'X' in appDict:
                    xrow = row
                    for xval in appDict['X']:
                        if xrow == row:
                            ws.write(xrow, 0, xval, self.header_format)
                        else:
                            ws.write(xrow, 0, xval, self.default)
                        xrow += 1
                # write a column of 'Y'
                if 'Y' in appDict:
                    yrow = row
                    for yval in appDict['Y']:
                        if yrow == row:
                            ws.write(yrow, 1, yval, self.header_format)
                        else:
                            ws.write(yrow, 1, yval, self.default)
                        yrow += 1
                # set column width at the end
                ws.set_column(0, 1, 20)
                print "Worksheet '%s' created." %(data)
            except:
                self.failedSheets.append(data)
                continue

    # input element tag returns the column index
    def getCol(self, tag):
        if tag.lower() in self.column:
            return self.column.index(tag.lower()) + 1 # column A is for tags
        else:
            return len(self.column) + 1

    # input an element and returns a boolean that indicates whether the element:
    # 1) has not text
    # 2) has at least a child node that has text
    # 3) has child that is not one of self.scalarUncertainty
    def childText(self, ele):
        hasCT = False # a flag for criteria 2)
        hasNS = False # a flag for criteria 3)
        if ele.text is not None:
            return False
        for child in ele.findall('./'):
            if child.tag not in self.scalarUncertainty:
                hasNS = True
            if child.text is not None:
                hasCT = True
            if hasCT and hasNS:
                return True
        return False

    # input an element and returns a boolean that indicates whether the element
    # has the same child nodes as the ScalarUncertainty type
    def typeSU(self, ele):
        childEles = ele.findall('./')
        # skip elements that have no child nodes
        if len(childEles) == 0:
            return False
        childTags = []
        for child in childEles:
            childTags.append(child.tag)
        childSet = set(childTags)
        sUSet = set(self.scalarUncertainty)
        # iff childSet is a subset of sUSet, return True
        if len(childSet.difference(sUSet)) == 0:
            return True
        return False

    # input an element and returns a boolean that indicates whether the element
    # has a child node that is the ScalarUncertainty type
    def childSU(self, ele):
        for child in ele.findall('./'):
            if self.typeSU(child):
                return True
        return False