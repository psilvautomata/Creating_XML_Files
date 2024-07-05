import xml.etree.ElementTree as ET
import pandas as pd
import os

#Naming the root and its elements
root = ET.Element("Your_Root_Name")
#Add many as you need
parent = ET.SubElement(root, "Your_Parent_Name")
child = ET.SubElement(root, "Your_Child_Name")
grandchild = ET.SubElement(root, "Your_Grandchild_Name")

#Values input (by user or using a loop to retrive its values)
root_val = "value"
parent_val = "value"
child_val = "value"
grandchild_val = "value"

file_path = 'Your/Path'
xml_tree = ET.ElementTree(root)
xml_file = os.path.join(file_path, "Your_File_Name.xml")
tree.write(file_path, encoding="utf-8", xml_declaration=True)
print(f"XML file '{xml_file}' created with sucess!")
