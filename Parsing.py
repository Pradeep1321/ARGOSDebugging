import xml.sax
import xmltodict
import collections
import openpyxl
import xlwings


tempFile = "C:\Share\Temp.xml"
#tempFile = "C:\\Users\\pnimmanapalli\\Downloads\\New folder\\run_at_config.xml"
ExcelFile = "C:\Share\Games.xlsx"

class GroupHandler(xml.sax.ContentHandler):
    def __init__(self):
        self.data = ''
    def startElement(self,name,attrs):
        self.current = name
        if attrs._attrs != {}:
            self.data= self.data+ "\n" + name + " = " + str(attrs._attrs)
    def characters(self, content):
        self.text = content

    def endElement(self, name):
        if self.current == name and self.text.strip() != '':
            self.data = self.data + "\n"+ name + " : " + self.text
    def datatobeReturned(self):
        return self.data

'''
handler = GroupHandler()
parser =   xml.sax.make_parser()
parser.setContentHandler(handler)
parser.parse(tempFile)
data = handler.datatobeReturned()
print("completed")
print(data)

'''
def ExcelUpdate(r,c,values):
    wb = xlwings.Book(ExcelFile)
    app = xlwings.apps.active
    ws = wb.sheets[0]
    oldval = ws.cells(r,c).value
    if oldval != None:
        ws.cells(r, c).value = oldval + "\n" + values
    else:
        ws.cells(r, c).value = values
    wb.save(ExcelFile)
    app.quit()
'''
def printDict(d):
    for k, v in d.items():
        if type(v) is collections.OrderedDict:
            printDict(v)
        elif type(v) is list:
            for j in v:
                if type(j) is collections.OrderedDict:
                    printDict(j)
        else:
            if k != None and v != None:
                #print("{0} : {1}".format(k, v))
                val = val + "\n" + str(k) + " : " + str(v)

handle = open(tempFile,"r")
content = handle.read()

dict = xmltodict.parse(content)

ExcelUpdate(2,5,data)
#value = printDict(dict)

'''

data = ''
File = "rr*.json"
tf = open(tempFile, "r")
dataList = tf.readlines()
if File == "build*.json":
    data = "".join([str(elem) for elem in dataList if elem.find("ModulePartNumber") != -1])
else:
    data = "".join([str(elem) for elem in dataList])
tf.close()
print(data)

ExcelUpdate(2,3,data)

