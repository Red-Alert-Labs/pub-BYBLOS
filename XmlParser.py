from io import StringIO
import xml.etree.ElementTree as ET


import docx
import xmlschema as xmlschema

from lxml import etree
def XMLPARSER():
    xsd_path=input('----> pls Enter your XSD(schema) file  path')
# open and read schema file
    with open(xsd_path, 'r') as schema_file:
        schema_to_check = schema_file.read()
        # print(schema_to_check)
    # open and read xml file
    xml_path = input('----> pls Enter your XML file  path')
    with open(xml_path, 'r') as xml_file:
        xml_to_check = xml_file.read()
        # print(xml_to_check)

        # xmlschema_check_error !!
        xmlschema_doc = etree.parse(StringIO(schema_to_check))
        xmlschema = etree.XMLSchema(xmlschema_doc)
        Syn_And_Val = False
        print(Syn_And_Val)
        # parse xml file !
        try:
            doc = etree.parse(StringIO(xml_to_check))
            print(doc.getroot())
            print('XML well formed, syntax ok.')

        # check for file IO error
        except IOError:
            print('Invalid File')

        # check for XML syntax errors
        except etree.XMLSyntaxError as err:
            print('XML Syntax Error, see error_syntax.log')
            with open('error_syntax.log', 'w') as error_log_file:
                error_log_file.write(str(err.error_log))
            quit()

        except:
            print('Unknown error, exiting.')
            quit()

    # validate XML file with schema !!
        try:
            xmlschema.assertValid(doc)
            Syn_And_Val = True
            print('XML valid, schema validation ok.')
            print(Syn_And_Val)

        except etree.DocumentInvalid as err:
            print('Schema validation error, see error_schema.log')
            with open('error_schema.log', 'w') as error_log_file:
                error_log_file.write(str(err.error_log))
            quit()

        except:
            print('Unknown error, exiting.')
            quit()
    # If xml file and xsd shema are valid !!!
        if Syn_And_Val==True:
            root = doc.getroot()
            # Create a new document. This command is equivalent to clicking 'new -> blank document' in MSWord
            document = docx.Document()
            for xx in range(len(root)):

                # Create table
                table = document.add_table(rows=7, cols=2)
                # Fill the first column of the table with the strings
                first_column = table.columns[0]
                first_column.cells[0].text = "Threat No"
                first_column.cells[1].text = "Threat"
                first_column.cells[2].text = "Threat description"
                first_column.cells[3].text = "Assets affected"
                first_column.cells[4].text = "Vulnerabilities"
                first_column.cells[5].text = "Threat Rating"
                first_column.cells[6].text = "countermeasures"
                # Fill the second column of the table with information from the xml file
                entry, vuln, countermeasure = "", "", ""
                second_column = table.columns[1]
                second_column.cells[0].text = root[xx][0].text
                second_column.cells[1].text = root[xx][1].text
                second_column.cells[2].text = root[xx][2].text
                second_column.cells[3].text = root[xx][3].text
                #we are under <vulnerabilities> wich =4 ; x parce all tags under it and get all <vul_name>&<vuln_desc> and fill them on column vulnerabilities of the word table
                for x in range(len(root[xx][4])):
                    vuln += root[xx][4][x][1].text + "\n"   #store all vul_name and description in vuln
                second_column.cells[4].text = vuln          #fill word table column vulnerabilities with vuln
                second_column.cells[5].text = root[xx][5].text
                for x in range(len(root[xx][4])):
                    countermeasure += "\n\n" + root[xx][4][x][3][1].text + "\n"

                    # This nested loop is to iterate through the <entry> tag to harvest the information and ttach to each countermeasure
                    for Z in range(len(root[xx][4][x][3][2])):
                        entry += root[xx][4][x][3][2][Z].text + "\n"
                        countermeasure += entry
                        entry = ""
                second_column.cells[6].text = countermeasure
                # This command breaks the tables into one table per page.
                document.add_page_break()
            document.save('test.docx')

XMLPARSER()