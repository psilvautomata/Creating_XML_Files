# $Creating$ $XML$ $Files$

<p align="justify">
Python code to create XML files using ElementTree, Pandas, os, and Datetime libraries.
</p>

<p align="justify">
XML files are an exceptional tool that simplifies data interaction with machines and facilitates readability for both humans and machines. They ensure the security and integrity of your data.
</p>

<p align = "justify">
The script contains the libraries: xml.etree.ElementTree for XML manipulation, pandas for handling Excel files, os for file path operations, and datetime for date and time operations.  
It starts reading data from an Excel file specified by the user and identifies the last filled row and the columns, extracting it values.  
The script then prepares for XML generation by iterating over the rows, starting from the second row (Empty cells in the Excel file are filled with empty strings to avoid printing zeros).  
Then creates an XML tree structure with a root element and several child elements.  
Finally, the script saves the XML file to a specified directory and prints a success message.
</p>
