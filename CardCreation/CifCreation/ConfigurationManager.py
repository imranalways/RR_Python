from xml.dom import minidom


def ConnectionString(tagName):
    xmldoc = minidom.parse(r"C:\RPA_Cardpro\CIF_Creation\Config.xml")
    # find the name element, if found return a list, get the first element
    name_element = xmldoc.getElementsByTagName(tagName)[0]
    # this will be a text node that contains the actual text
    text_node = name_element.childNodes[0]
    return text_node.data


def RobotData(tagName):
    xmldoc = minidom.parse(r"C:\RPA_Cardpro\CIF_Creation\App.xml")
    # find the name element, if found return a list, get the first element
    name_element = xmldoc.getElementsByTagName(tagName)[0]
    # this will be a text node that contains the actual text
    text_node = name_element.childNodes[0]
    return text_node.data
