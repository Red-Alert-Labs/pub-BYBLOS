import xml.etree.ElementTree as ET
from docx import Document
def XmlToDoc(path):
    # Parse the xml file to get its root
    tree = ET.parse(path)
    root = tree.getroot()
    # Create a new document. This command is equivalent to clicking 'new -> blank document' in MSWord
    document = Document()
    # Loop through each child node of root (in this case it is the threats).
    # Within the loop, a table is created in the word file and populated with info from the xml file
    for xx in range(len(root)):

        #Create table
        table = document.add_table(rows=7, cols=2,)
        table.style = 'TableGrid'
        #Fill the first column of the table with the strings
        first_column = table.columns[0]
        first_column.cells[0].text = "Threat No"
        first_column.cells[1].text = "Threat"
        first_column.cells[2].text = "Threat description"
        first_column.cells[3].text = "Assets affected"
        first_column.cells[4].text = "Vulnerabilities"
        first_column.cells[5].text = "Threat Rating"
        first_column.cells[6].text = "countermeasures"
        #Fill the second column of the table with information from the xml file
        entry,vuln,countermeasure = "", "",""
        second_column = table.columns[1]
        second_column.cells[0].text = root[xx][0].text
        second_column.cells[1].text = root[xx][1].text
        second_column.cells[2].text = root[xx][2].text
        second_column.cells[3].text = root[xx][3].text
        for x in range(len(root[xx][4])):
            vuln += root[xx][4][x][1].text + "\n"
        second_column.cells[4].text = vuln
        second_column.cells[5].text = root[xx][5].text
        for x in range(len(root[xx][4])):
            countermeasure += "\n\n" + root[xx][4][x][3][1].text + "\n"

            #This nested loop is to iterate through the <entry> tag to harvest the information and ttach to each countermeasure
            for Z in range(len(root[xx][4][x][3][2])):
                entry += root[xx][4][x][3][2][Z].text + "\n"
                countermeasure += entry
                entry = ""
        second_column.cells[6].text = countermeasure
        #This command breaks the tables into one table per page.
        document.add_page_break()
    document.save('test.docx')



XmlToDoc(r'C:\Users\RAL-lab\PycharmProjects\Firefox\TestEnv\sample2.xml')