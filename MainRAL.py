# import ExeclToXml
import XmlParser
# import xmltoDoc2



def main():
    print('################### RAL DATA FLOW EXCHANGE ########################\n')
    print('[+]______FOR VALIDATE YOUR XML/XSD FILE PRESS_________[1]\n ')
    print('[+]______FOR CREATE YOUR XML FROM EXCEL FILE PRESS____[2] \n')
    print('[+]______FOR CREATE YOUR WORD FROM XML FILE PRESS_____[3] \n')
    print('################### ###################### ########################\n')
    answer=input ('|------> Pls Select your Answer(1..4) : ')

    if answer == 1 :
        XmlParser.XMLPARSER()
    else:
        print('bye')


main()
