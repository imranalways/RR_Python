from xml.dom import minidom
import pandas as pd


# def filexml(tagname):
#
#     xmldoc = minidom.parse(r"C:\Users\imran.hossain\Downloads\ALL Files\PYTHON\Risk_Resister_Python\Xpath.xml")
#     # find the name element, if found return a list, get the first element
#     name_element = xmldoc.getElementsByTagName(tagname)[0]
#     # this will be a text node that contains the actual text
#     text_node = name_element.childNodes[0]
#     return text_node.data


def fileXpathxlsx():
    df = pd.read_excel(r'C:\Users\imran.hossain\Downloads\ALL Files\PYTHON\Risk_Resister_Python\Xpath.xlsx')
    return(df)
def fileDataxlsx():
    df = pd.read_excel(r'C:\Users\imran.hossain\Downloads\ALL Files\PYTHON\Risk_Resister_Python\data.xlsx')
    return(df)