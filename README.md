PyPages
=======

Python tool for manipulating pages documents

Example load commands:
import pages,os

myDoc = pages.PagesDoc(filePath="/tmp/tablepages.pages")

myDoc.unPackFile()

## myDoc.packUpFile("/tmp/myPages.pages")

##newParagraph = myDoc.newParagraphNode(text="hellothere")

##myDoc.addElementToDOM(newParagraph)

#myDoc.writeXML()

#myDoc.packFile("/tmp/newpages.pages")

#myDoc.writeXML("/tmp/cleanpages1.xml",pretty=True)



## Create a table

tableData = []

tableData.append({"col1":"row1cell1","col2":"row1cell2","col3":"row1cell3"})

tableData.append({"col1":"row2cell1","col2":"row2cell2","col4":"row2cell4"})

tableData.append({"col1":"row3cell1","col5":"row3cell5","col3":"row3cell3"})



tableHeaders = []

tableHeaders.append({"displayName":"Col 1","dataKey":"col1","width" : 150 })

tableHeaders.append({"displayName":"Col 2","dataKey":"col2","width" : 75 })

tableHeaders.append({"displayName":"Col 3","dataKey":"col3"})

tableHeaders.append({"displayName":"Col 4","dataKey":"col4"})

tableHeaders.append({"displayName":"Col 5","dataKey":"col5"})



myDoc.addNewTableToDOM(tableData=tableData,tableHeaders=tableHeaders)

tableHeaders2 = tableHeaders[0:2]

myDoc.addNewTableToDOM(tableData=tableData,tableHeaders=tableHeaders2)

newParagraph = myDoc.newParagraphNode(text="test text (this should be a header)")

myDoc.addElementToDOM(newParagraph)

tableData.pop()

myDoc.addNewTableToDOM(tableData=tableData,tableHeaders=tableHeaders2,tableWidth=200)



## Create a table

tableData3 = []

tableHeaders3 = tableHeaders[0:4]

tableHeaders3[0]["cellStyle"] = "SFTCellStyle-9"

tableHeaders3[3]["cellStyle"] = "SFTCellStyle-4"

tableHeaders3[0]["width"] = 20

tableData3.append({"col1":"row1cell1","col2":"row1cell2","col3":"row1cell3","col4":"row1cell4"})

tableData3.append({"col1":"row2cell1","col2":"row2cell2","col4":"row2cell4"})

tableData3.append({"col1":"row3cell1","col4":"row3cell4","col3":"row3cell3"})

tableData3.append({"col1":"row4cell1","col2":"row4cell2","col3":"row4cell3","col4":"row4cell4"})

tableData3.append({"col1":"row5cell1","col2":"row5cell2","col3":"row5cell3","col4":"row5cell4"})

tableData3.append({"col1":"row6cell1","col2":"row6cell2","col3":"row6cell3","col4":"row6cell4"})

myDoc.addNewTableToDOM(tableData=tableData3,tableHeaders=tableHeaders3)



## Create a table and add it to the DOM (use addNewTableToDOM)

##myTable = myDoc.newTableNode(tableData=tableData,tableHeaders=tableHeaders)

##myDoc.addAttachmentToDOM(myTable,kind="tabular-attachment",id="SFTTableAttachment-1")

myDoc.writeXML("./table2generated.xml",pretty=True)

##myDoc.writeXML("/tmp/cleanpages1.xml",pretty=True)

myDoc.writeXML()

myDoc.packFile("/tmp/newpages1.pages")
