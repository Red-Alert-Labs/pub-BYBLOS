from xml.etree.ElementTree import Element, SubElement, Comment, tostring
import xlrd as pd
import lxml.etree
import xml.etree.cElementTree as ET

#Open xlsx file
book = pd.open_workbook(r'C:\Users\RAL-lab\Desktop\data.xlsx')
#pick the first xlsx sheet
first_sheet = book.sheet_by_index(0)
# take the second sheet
second_sheet = book.sheet_by_index(1)

# define the root element for the XML
threat = Element('threat')

# define subelement of threat named vulnerabilities
vulnerabilities = SubElement(threat, 'Vulnerabilities')

# the tag vulnerabilities contain * tag vulnerability
for row in range(3, first_sheet.nrows):

    # the tag vulnerability is subelement of vulnerabilities in the XML file
    vulnerability = SubElement(vulnerabilities, "vulnerability")
    #the subelement id of vulnerability tag has a value wich is the value of the first_sheet (row)column(1)
    vuln_id=SubElement(vulnerability,'vuln_id')
    vuln_id.text=first_sheet.row_values(row)[1]

    vuln_name=SubElement(vulnerability,'vuln_name')
    vuln_name.text=first_sheet.row_values(row)[2]

    vuln_description=SubElement(vulnerability,'vuln_description')
    vuln_description.text=first_sheet.row_values(row)[3]

    countermesure=SubElement(vulnerability,'countermesure')
    cm_id=SubElement(countermesure,'cm_id')
    cm_id.text=first_sheet.row_values(row)[6]

    cm_name=SubElement(countermesure,'cm_name')
    cm_name.text=first_sheet.row_values(row)[5]

    entry=SubElement(countermesure,'entry')
    # the two sheet are related by the countermesure so we need to memorise the last contermesureid to check other information related to it in the other sheet
    CM_Entry=cm_id.text
    for rows in range(3, second_sheet.nrows):
        # check if the cm_id of the first_sheet is the same raw[column(0)]of the second sheet continue to make subelement and extract the other information and put it in XML file
        if second_sheet.row_values(rows)[0]==CM_Entry:
            cm_entry=SubElement(entry,'cm_entry')
            cm_entry.set('id',second_sheet.row_values(rows)[1])
            cm_entry.text=second_sheet.row_values(rows)[2]
    # print all the tags starting from the root element
    print(tostring(threat))
    # API for parsing and creating XML data.
    tree = ET.ElementTree(threat)
    tree.write("test.xml")


