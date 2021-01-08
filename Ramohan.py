import xml.etree.ElementTree as ET

File1 = "C:\\Users\\pnimmanapalli\\Downloads\\New folder\\run_at_config.xml"
File2 = "C:\\Users\\pnimmanapalli\\Downloads\\New folder\\blank_config.xml"

def compare(tree1,tree2):
    for i in range(len(tree1)):
        chld1 = list(tree1[i])
        chld2 = list(tree2[i])
        if len(chld1) == len(chld2):
            for j in range(len(chld1)):
                if chld1[j].tag == chld2[j].tag:
                    chld2[j].text = chld1[j].text

def until(tree,item):
    for el in tree.iter():
        if el.tag == item:
            return el
    return None

tree1 = ET.parse(File1)
tree2 = ET.parse(File2)
item1 = "hudson.model.ParametersDefinitionProperty"
item2 = "properties"
compareditem = "parameterDefinitions"
tr1 = until(tree1,compareditem)
tr2 = until(tree2,compareditem)
if not tr2:
    tr2 = until(tree2, item2)
    tr2.append(tr1)
    tree2.write(File2)
else:
    lst1 = list(tr1)
    lst2 = list(tr2)
    if len(lst1) == len(lst2):
        compare(lst1,lst2)
    print(lst1)
    print(lst2)
    tree2.write(File2)
