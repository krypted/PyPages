#!/usr/bin/env Python


################################
##
##  Pages
##  A Pages XML file manipulation tool 
##  
##  This class provides the primary interface for manipulation of an Apple 
##  Pages XML document.
## 
#############################################################

import sys,os,shutil,subprocess
import codecs
import uuid
import re,datetime,time,tempfile,copy
import zipfile
from xml.dom import minidom

class BaseObject:
  """Our base object. Provides basic error logging capabilities"""
  
  ## Logging vars
  log = []
  lastError = ""
  lastMSG = ""
  debug = False
  isError = False
  logOffset = 1
  printLogs = True
  
  timezoneOffset = 5
  
  def __init__(self):
    """Our construct"""
    self.log = []
    self.lastError = ""
    self.lastMSG = ""
    self.debug = True
    self.isError = False
    self.logOffset = 1
    self.printLogs = True
    
  def logger(self, logMSG, logLevel="normal"):
    """(very) Basic Logging Function, we'll probably migrate to msg module"""
    if logLevel == "error" or logLevel == "normal" or self.debug:
        i = 0
        headerText = ""
        while i < self.logOffset:
          headerText+="    "
          i+=1
        if logLevel.lower() == "error":
          headerText+="ERROR: "
        if self.printLogs:
          print "%s%s: %s" % (headerText,self.__class__.__name__,logMSG)
          sys.stdout.flush()
        self.lastMSG = logMSG
    if logLevel == "error":
        self.lastError = logMSG
    self.log.append({"logLevel" : logLevel, "logMSG" : logMSG})
    
  def logs(self,logLevel=""):
    """Returns an array of logs matching logLevel"""
    returnedLogs = []
    logs = self.log
    for log in logs:
      if logLevel and logLevel.lower() == log["logLevel"].lower():
        returnedLogs.append(log)


class PagesDoc(BaseObject):
  
  filePath = ""           ## Absolute path to pages file
  unPackedPath = ""       ## Absolute path to unpacked directory
  outFilePath = ""        ## Absolute path to our destination file
  
  
  ## Styling information
  
  defaultParagraphStyle = "paragraph-style-32"
  defaultTableStyle = "SFTTableStyle"
  
  nextTableIterator = 0
  
  
  DOM = ""                ## Our document object model (xml.dom.minidom)
  
  def __init__(self,filePath=""):
    """Our construct, accepts a single argument, a path to our pages file"""

    ## Instantiate our member vars
    self.filePath = ""
    self.outFilePath = ""
    self.unPackedPath = ""
    ## todo: determine nextTableIterator from pages doc
    nextTableIterator = 0

    ## Styling
    self.defaultParagraphStyle = "paragraph-style-32"
    self.defaultTableStyle = "SFTTableStyle"

    ## Debug Setting
    self.debug = True
    
    ## If we were passed a filePath, set it
    if filePath:
      self.setFile(filePath)
    
  def setFile(self,filePath):
    """Sets the operating pages file to file found at filePath"""
    if not os.path.isfile(filePath):
      self.logger("Could not find file at path: '%s'!" % filePath,"error")
      return False
    self.logger("Initing with filePath: '%s'" % filePath,"debug")
    self.filePath = os.path.abspath(u"%s" % filePath)
    
  def setOutFile(self,outFilePath):
    """Sets our destination file"""
    if not os.path.exists(os.path.dirname(outFilePath)):
      self.logger("Could not set outfile path to '%s'," % outFilePath
            + "parent directory does not exist!", "error")
      return False
    self.outFilePath = os.path.abspath(u"%s" % outFilePath)
    
    return False

  def unPackFile(self,filePath=""):
    """Unzips the specified pages file to a temporary directory"""
    
    tempPathPrefix = "/private/tmp/pypages"
    
    if not filePath:
      if self.filePath:
        filePath = self.filePath
      else:
        self.logger("Could not unpack file, no pages file specified!","error")
        return False
        
    self.logger("Unpacking file: '%s'" % filePath,"debug")
    
    ## Make sure we have a good file
    if not os.path.isfile(filePath):
      self.logger("Could not unpack file, no such file exists at '%s'!","error")
      return False
      
    ## Create our temporary directory
    try:
      tempDir = tempfile.mkdtemp(prefix=tempPathPrefix)
    except:
      self.logger("Could not create temporary directory at path: '%s'," \
                  " cannot continue!" % filePath,"error")
    
    ## Attempt to unzip our file
    self.logger("Attempting to unzip file to:'%s'" % tempDir,"debug")
    try:
      myZipFile = zipfile.ZipFile(filePath,"r")
      myZipFile.extractall(tempDir)
    except:
      self.logger("An error occured extracting the zip file:'%s'" % filePath,"error")
      if self.debug:
        raise
      return False
    
    self.logger("Successfully unpacked file to:'%s'" % tempDir,"debug")
    
    ## Update our member var
    self.unPackedPath = tempDir

  def packFile(self,outFilePath="",overwriteExistingFiles="replace"):
    """Unzips the unpacked pages file to outFilePath"""
    
    if not outFilePath:
      if self.outFilePath:
        filePath = self.outFilePath
      else:
        self.logger("Could not pack file, no destination specified!","error")
        return False
    if not self.unPackedPath:
      self.logger("Cannot pack file, no unpacked directory is set!","error")
      return False
      
    unPackedPath = self.unPackedPath
    
    self.logger("Packing dir:'%s' to file: '%s'" % (unPackedPath,outFilePath),"debug")
    
    ## Make sure we have a good file
    if os.path.isfile(outFilePath):
      if overwriteExistingFiles.lower() == "ask":
        print "File already exists at path: '%s'" % outFilePath
        confirmation = ""
        while (confirmation is not "n" and confirmation is not "y"
        and confirmation is not "no" and confirmation is not "yes"):
          confirmation = raw_input("overwrite existing file? y/n: ");
        if confirmation is "no" or confirmation is "n":
          return False
        ## delete the existing file
        os.remove(outFilePath)
      elif overwriteExistingFiles.lower() == "iterate":
        iterCount = 1
        fileName = os.path.basename(outFilePath)
        fileDir = os.path.dirname(outFilePath)
        fileBase,fileExtension = os.path.splitext(outFilePath)
        while os.path.exists(outFilePath):
          outFilePath = os.path.join(fileDir,"%s-%s%s" % (fileBase,iterCount,fileExtension))
          iterCount += 1
      elif overwriteExistingFiles.lower() == "replace":
        ## delete the existing file
        self.logger("Overwriting file at path: '%s'" % outFilePath,"debug")
        os.remove(outFilePath)
      else:
        self.logger("Could not pack file, file already exists at '%s'!" % outFilePath,"error")
        return False
    
    ## Attempt to unzip our file
    try:
      myZipFile = zipfile.ZipFile(outFilePath,"w")
      myZipFile.write(os.path.join(unPackedPath,"QuickLook/Preview.pdf"),"QuickLook/Preview.pdf")
      myZipFile.write(os.path.join(unPackedPath,"QuickLook/Thumbnail.jpg"),"QuickLook/Preview.pdf")
      myZipFile.write(os.path.join(unPackedPath,"buildVersionHistory.plist"),"buildVersionHistory.plist")
      myZipFile.write(os.path.join(unPackedPath,"index.xml"),"index.xml")
    except:
      self.logger("An error occured extracting the zip file:'%s'" % filePath,"error")
      if self.debug:
        raise
      return False
    
    self.logger("Successfully packed file to:'%s'" % outFilePath,"debug")
    
  def buildXML(self):
    """Writes out the Pages XML file"""
    
    ## Read in our XML file, first verify that we have a file to read in
    if self.unPackedPath:
      filePath = os.path.join(self.unPackedPath,"index.xml")
    else:
      self.logger("Could not build XML! No unpacked Pages file found!","error")
      return False
    
    if not os.path.isfile(filePath):
      self.logger("Could not build XML! File not found: '%s'" % filePath,"error")
      return False
    
    try:
      myDOM = minidom.parse(filePath).documentElement
    except:
      self.logger("An unknown error occured reading XML file: '%s'" % filePath)
      if self.debug:
        raise
      return False
    
    ## Inject our data
    
    ## Save our new dom
    self.DOM = myDOM
    
    return self.DOM
    
  def writeXML(self,filePath="",overwriteExistingFiles="replace",pretty=False):
    """Writes our XML to specified filepath, can output 'pretty' if specified
    (line breaks and indentation)"""
    
    ## Make sure we have a filePath
    if not filePath:
      if self.unPackedPath:
        filePath = os.path.join(self.unPackedPath,"index.xml")
      else:
        self.logger("Could not pack file, no destination specified!","error")
        return False
    
    ## Get our DOM, build it if neccessary
    if self.DOM:
      DOM = self.DOM
    else:
      DOM = self.buildXML()
      
    if not DOM:
      self.logger("Could not write XML file! DOM could not be built!","error")
      return False
    
    ## Check to see if our path exists. 
    if os.path.isfile(filePath):
      if overwriteExistingFiles.lower() == "ask":
        print "File already exists at path: '%s'" % opt[1]
        confirmation = ""
        while (confirmation is not "n" and confirmation is not "y"
        and confirmation is not "no" and confirmation is not "yes"):
                confirmation = raw_input("overwrite existing file? y/n: ");
        if confirmation is "no" or confirmation is "n":
            self.logger("exiting...\n");
            return 2
            
        ## delete the existing file
        os.remove(filePath)
   
      elif overwriteExistingFiles.lower() == "iterate":
        iterCount = 1
        fileName = os.path.basename(filePath)
        fileDir = os.path.dirname(filePath)
        fileBase,fileExtension = os.path.splitext(fileName)
        while os.path.exists(filePath):
          filePath = os.path.join(fileDir,"%s-%s%s" % (fileBase,iterCount,fileExtension))
          iterCount += 1
      elif overwriteExistingFiles.lower() == "replace":
        ## delete the existing file
        self.logger("Overwriting file at path: '%s'" % filePath,"debug")
        os.remove(filePath)
  
    elif os.path.exists(filePath):
      self.logger("Non-file object resides at: '%s', cannot continue!","error")
      return False
    elif not os.path.exists(os.path.dirname(filePath)):
      self.logger("Could not write XML file, parent directory: '%s' does not exist!" % os.path.dirname(filePath),"error")
      return False    
      
    ## Write out our XML file
    wasError = False
    theFile = codecs.open(filePath,"w","utf-8")
    try:
      if not pretty: 
        DOM.writexml(theFile)
      else:
        xml = DOM.toprettyxml()
        theFile.write(xml)
    except:
      wasError = True
      self.logger("An unknown error occured writing XML!","error")
      if self.debug:
        raise
    
    theFile.close()
    
    if not wasError:
      return True
    else:
      return False
  
  def nodeFromPath(self,path,DOM=""):
    """Returns a node from a filesystem-like path 
    (/sf:text-storage/sf:attachments)"""
    
    ## resolve our DOM
    if not DOM:
      if self.DOM:
        DOM = self.DOM
      else:
        DOM = self.buildXML()

    if not DOM:
      self.logger("No DOM loaded, cannot continue!","error")
      return False
        
    ## Explode our path
    pathArray = path.split("/")
    if len(pathArray) < 1 or len(pathArray) == 1 and not pathArray[0]:
      self.logger("Provided path has no elements!","error")
      return False
    
    myNode = ""
    currentParent = DOM
    element = ""
    ## Iterate through our pathArray looking for each tier,
    ## traversing the DOM as we go.
    for pathElement in pathArray:
      if not pathElement:
        continue
      didFindElement = False
      for element in currentParent.childNodes:
        if element.nodeName == pathElement:
          didFindElement = True
          break;
      
      if didFindElement:
        currentParent = element
        myNode = currentParent
      else:
        myNode = ""
        self.logger("Failed to find node: '%s'" % pathElement,"error")
        return False
      
    if not myNode:
      self.logger("Failed to find node with path: '%s'" % path,"error")
      return False

    return myNode
    
  def addAttachmentToDOM(self,newElement,kind,id):
    """Adds an attachment to the DOM, specify kind (i.e. "tabular-attachment")
    and element id attribute (i.e. "SFTTableAttachment-0")"""
    ## See if we have a DOM
    if self.DOM:
      DOM = self.DOM
    else:
      DOM = self.buildXML()
      
    if not DOM:
      self.logger("No DOM loaded, cannot continue!","error")
      return False
      
    xmlDoc = minidom.Document()

    ## Get our attachment node
    attachmentsNode = self.nodeFromPath("/sf:text-storage/sf:attachments")
    textBodyNode = self.nodeFromPath("/sf:text-storage/sf:text-body")
    
    if not attachmentsNode:
      attachmentsNode = xmlDoc.createElement("sf:attachments")
      textStorageNode = self.nodeFromPath("/sf:text-storage")
      if textStorageNode:
        textStorageNode.insertBefore(attachmentsNode,textBodyNode)
      else:
        self.logger("Could not process DOM: sf:text-storage node not found!","error")
        return False
    
    attachmentElement = xmlDoc.createElement("sf:attachment")
    attachmentElement.setAttribute("sf:kind",kind)
    attachmentElement.setAttribute("sfa:ID",id)
    attachmentElement.setAttribute("sfa:sfclass","")
     
    ## Make sure our ID is unique
    for attachment in attachmentsNode.childNodes:
      if attachment.attributes["sfa:ID"] == id:
        self.logger("Could not add attachment: '%s' to DOM, attachment with ID "
          + "already exists!","error")
        return False
        
    try:
      attachmentElement.appendChild(newElement)
      positionElement = xmlDoc.createElement("sf:position")
      positionElement.setAttribute("sfa:x","72")
      positionElement.setAttribute("sfa:y","219")
      attachmentElement.appendChild(positionElement)
      attachmentsNode.appendChild(attachmentElement)

      
    except:
      self.logger("An unknown error occured adding attachment to the DOM!","error")
      return False
    
    
    ## Next, add our attachment reference into the main DOM tree
    
    textBodyNode = self.nodeFromPath("/sf:text-storage/sf:text-body")
    if not textBodyNode:
      self.logger("Could not process DOM: sf:text-body node not found!","error")
      return False      
    
    ## todo: doesn't look like we need these, need to figure out what they're for
    selectionStartElement = xmlDoc.createElement("sf:selection-start")
    selectionEndElement = xmlDoc.createElement("sf:selection-end")
    
    
    attachmentLinkElement = xmlDoc.createElement("sf:attachment-ref")
    attachmentLinkElement.setAttribute("sfa:IDREF",id)
    attachmentLinkElement.setAttribute("sf:kind",kind)
    
    paragraphElement = self.newParagraphNode(newElement=attachmentLinkElement)
    
    self.addElementToDOM(selectionStartElement)
    paragraphElement.appendChild(selectionEndElement)
    self.addElementToDOM(paragraphElement)
    
    return True
  
  def addNewTableToDOM(self,tableData,tableHeaders="",style="",tableWidth=467,tableHeight=75):
    """Creates a tableNode and adds it to the DOM"""
    
    if not style:
      style = self.defaultTableStyle
    
    theTable = self.newTableNode(tableData=tableData,tableHeaders=tableHeaders,style=style,tableWidth=tableWidth,tableHeight=tableHeight)
    if not theTable:
      self.logger("Could not create table from passed data!","error")
      return False
    self.nextTableIterator += 1
    self.addAttachmentToDOM(theTable,kind="tabular-attachment",id="SFTTableAttachment-%s" % self.nextTableIterator)
    
    ## Create our stylesheet node
    stylesNode = self.nodeFromPath("/sl:stylesheet/sf:anon-styles")
    ## Find our tablestyle node todo:this is static and hacky, fix it
    didDuplicateStyle = False
    nodeIDToClone = "%s-1" % style
    for childElement in stylesNode.childNodes:
      self.logger("Searching for sf:tabular-style node. current node:'%s'" % childElement.nodeName,"debug")
      try:
        childElementID = childElement.attributes["sfa:ID"].value
        if childElement.nodeName == "sf:tabular-style":
            self.logger("Searching for node: '%s' to clone, current node:'%s'" % (nodeIDToClone,childElementID),"debug")
            if childElementID == nodeIDToClone:
              self.logger("Cloning node:'%s' new style:'%s-%s'" % (nodeIDToClone,style,self.nextTableIterator),"debug")
              newStyleNode = childElement.cloneNode(True)
              newStyleNode.setAttribute("sfa:ID","%s-%s" % (style,self.nextTableIterator))
              stylesNode.appendChild(newStyleNode)
              didDuplicateStyle = True
              break
      except:
        self.logger("cloneNode: Skipping node, no ID found","debug")
    if not didDuplicateStyle:
      self.logger("Could not clone node '%s'!" % nodeIDToClone,"error")
      return False
    
    return theTable

    
  def addElementToDOM(self,newElement):
    """Returns a minidom element for our storage node"""
    
    ## See if we have a DOM
    if self.DOM:
      DOM = self.DOM
    else:
      DOM = self.buildXML()
      
    if not DOM:
      self.logger("No DOM loaded, cannot continue!","error")
      return False

    ## Get our text-storage node
    textStorageNode = ""
    for element in DOM.childNodes:
      if element.nodeName == "sf:text-storage":
        textStorageNode = element
        break;
    if not textStorageNode:
      self.logger("Could not process DOM: sf:text-storage node not found!","error")
      return False
  
    ## Get our attachment and text-body nodes
    attachmentNode = ""
    textBodyNode = ""
    for element in textStorageNode.childNodes:
      if element.nodeName == "sf:attachments":
        attachmentNode = element
      elif element.nodeName == "sf:text-body":
        textBodyNode = element
    
    if not attachmentNode or not textBodyNode:
      self.logger("Could not process DOM: sf:attachments or sf:text-body node not found!","error")
      return False
    
    ## Get our sf:section node
    sectionNode = ""
    for element in textBodyNode.childNodes:
      if element.nodeName == "sf:section":
        sectionNode = element
    if not sectionNode:
      self.logger("Could not process DOM: sf:section node not found!","error")
      return False      
    
    ## Get our layout node
    layoutNode = ""
    for element in sectionNode.childNodes:
      if element.nodeName == "sf:layout":
        layoutNode = element
        
    if not layoutNode:
      self.logger("Could not process DOM: sf:layout node not found!","error")
      return False      
    
    ## Finally, append our new node
    layoutNode.appendChild(newElement)
    
      
  def newTextNode(self,text,style=""):
    """Returns a DOM object for a basic text node (used for inline formatting)"""
    xmlDoc = minidom.Document()
    newNode = xmlDoc.createTextNode(u"%s" % text)
    return newNode
  
  def newLineBreakNode(self,):
    """Returns a DOM object for a basic line break <br>"""
    xmlDoc = minidom.Document()
    lineBreakNode = xmlDoc.createElement("sf:br")
    return lineBreakNode
  
  def newParagraphNode(self,newElement="",text="",style=""):
    """Returns a DOM object for a basic paragraph with text"""
    
    if not style:
      style = self.defaultParagraphStyle
    
    xmlDoc = minidom.Document()
    paragraphElement = xmlDoc.createElement("sf:p")
    paragraphElement.setAttribute("sf:style",style)
    
    if text:
      paragraphElement.appendChild(self.newTextNode(text))

    if newElement:
      paragraphElement.appendChild(newElement)
      
    paragraphElement.appendChild(self.newLineBreakNode())
    
    return paragraphElement

  def _newTableStyle(self,style=""):
    """Creates a DOM object for a new table style"""
    xmlDoc = minidom.Document()
    tablularStyle = xmlDoc.createElement("sf:tabular-style")
    ## todo: need to implement this???
    return False
    
    

  def newTableNode(self,tableData,tableHeaders="",tableName="",style="",tableWidth=467,tableHeight=75):
    """Returns a DOM object for a table accepts a dictionary of tableHeaders:
    {"hostname":{"displayName":"Name","dataKey":"hostname",width:100},
    "ip":{"displayName":"Primary IP Address","dataKey":"en0ip"}
    }
    
    and TableData:
    [{"hostname":"server1","en0ip":"10.0.2.1"}]
    
    You can also specify the style of the table
    """
    
    sizesLocked = "true"
    #tableWidth = "311.33334350585938"
    ##tableWidth = 467
    ##tableHeight = 75
    minCellWidth = 20
    minCellHeight = 25
    
    ## build our data headers
    standardHeaderStyleName = "SFTCellStyle-6"
    standardCellStyleName = "SFTCellStyle-2"
    
    cellPaddingX = 5
    cellPaddingY = 1
    
    ## debug:
    ##minCellWidth = 45.337891
  
    posX = "0"
    posY = "0"
    
    if not style:
      style = self.defaultTableStyle
      
    if not style:
      style = "SFTTableStyle"
    
    ##debug
    style = "SFTTableStyle-2"
    
    try:
      if len(tableData) == 0:
        self.logger("No table data provided, cannot continue!","error")
        return False
    except:
      self.logger("Cannot interpret tableData! cannot continue!")
      return False
    
    ## Figure out some basic info
    if tableHeaders:
      numColumns = len(tableHeaders)
      numHeaderRows = 1
    else:
      numColumns = 0
      numHeaderRows= 0
      for row in tableData:
        if len(row) > numColumns:
          numColumns = len(row)
      
    numRows = len(tableData)
    if tableHeaders:
      numRows += 1
    ## Todo: make numHeaderColumns a passable option
    numHeaderColumns = 0
    
    numCells = numRows * numColumns
    
    ## Create our table DOM
    xmlDoc = minidom.Document()
    tableMaster = xmlDoc.createElement("sf:tabular-info")
    
    ## Our Geometry block
    geometryElement = xmlDoc.createElement("sf:geometry")
    geometryElement.setAttribute("sf:sizesLocked",sizesLocked)
    
    naturalSizeElement = xmlDoc.createElement("sf:naturalSize")
    naturalSizeElement.setAttribute("sfa:h","%s" % tableHeight)
    naturalSizeElement.setAttribute("sfa:w","%s" % tableWidth)
    geometryElement.appendChild(naturalSizeElement)
    
    sizeElement = xmlDoc.createElement("sf:size")
    sizeElement.setAttribute("sfa:h","%s" % tableHeight)
    sizeElement.setAttribute("sfa:w","%s" % tableWidth)
    geometryElement.appendChild(sizeElement)
        
    positionElement = xmlDoc.createElement("sf:position")
    positionElement.setAttribute("sfa:x","%s" % posX)
    positionElement.setAttribute("sfa:y","%s" % posY)
    geometryElement.appendChild(positionElement)
    
    ## Append our geometry block to our table element.
    tableMaster.appendChild(geometryElement)
    
    ## Create our style block
    styleElement = xmlDoc.createElement("sf:style")
    tableStyleElement = xmlDoc.createElement("sf:tabular-style-ref")
    tableStyleElement.setAttribute("sfa:IDREF","%s" % style)
    styleElement.appendChild(tableStyleElement)
    
    ## Append it to our master
    tableMaster.appendChild(styleElement)
    
    ## Create our table model
    tableModelElement = xmlDoc.createElement("sf:tabular-model")
    tableModelElement.setAttribute("sf:grouping-enabled","false")
    tableModelElement.setAttribute("sf:header-columns-frozen","false")
    tableModelElement.setAttribute("sf:header-rows-frozen","false")

    ## Create a unique ID for our table
    uuidString = "%s" % uuid.uuid1()
    uuidString = uuidString.upper().replace("-","")
    tableModelElement.setAttribute("sf:id","%s" % uuidString)
    if tableName:
      tableModelElement.setAttribute("sf:name","%s" % tableName)
      tableModelElement.setAttribute("sf:name-is-visible","true")
    else:
      tableModelElement.setAttribute("sf:name","Unnamed Table")
      tableModelElement.setAttribute("sf:name-is-visible","false")   
    
    tableModelElement.setAttribute("sf:num-footer-rows","0")   
    tableModelElement.setAttribute("sf:num-header-columns", "%s" % numHeaderColumns)   
    tableModelElement.setAttribute("sf:num-header-rows","%s" % numHeaderRows)   
        
    ## Our tabular style block
    tableStyleElement = xmlDoc.createElement("sf:tabular-style-ref")
    tableStyleElement.setAttribute("sfa:IDREF","%s" % style)
    tableModelElement.appendChild(tableStyleElement)
    
    ## Create our grid
    gridElement = xmlDoc.createElement("sf:grid")
    gridElement.setAttribute("sf:hiddennumcols","0")
    gridElement.setAttribute("sf:hiddennumrows","0")
    gridElement.setAttribute("sf:numcols","%s" % numColumns)
    gridElement.setAttribute("sf:numrows","%s" % numRows)
    ## Not sure what these are for
    gridElement.setAttribute("sf:ncc","true")
    gridElement.setAttribute("sf:ocnt","%s" % numCells)

    ## Our columns
    columnsElement = xmlDoc.createElement("sf:columns")
    columnsElement.setAttribute("sf:count", "%s" % numColumns)
    
    ## Determine our column widths 
    if tableHeaders:
      ## Determine appropriate width of each column 
      ## try to respect widths set in tableHeaders
      currentUsedWidth = 0        ## Tracks total used space for cols
      preferredRequestedWidth = 0 ## Tracks requested space to fixed width cols
      preferredUsedWidth = 0      ## Tracks provisioned space to fixed width cols
      columnMinWidth = minCellWidth    ## Soft setting
      autoColumns = []
      fixedColumns = []
      count = 0
      ## Iterate through each header,allocate widths
      for header in tableHeaders:
        headerIndex = count
        count += 1
        if "width" in header and header["width"]:
          autoWidth = False
          headerWidth = header["width"]
          fixedColumns.append(header)
        else:
          autoWidth = True
          headerWidth = columnMinWidth
          header["width"] = headerWidth
          autoColumns.append(header)
        
        ## Set our starting actualWidth value
        header["actualWidth"] = headerWidth
        
        ## Check to see if our table can add the column without overflowing
        if not headerWidth + currentUsedWidth > tableWidth:
          ## Here if we're in bounds
          if not autoWidth:
            preferredRequestedWidth += headerWidth
            preferredUsedWidth += headerWidth
          currentUsedWidth += headerWidth          

        else:
          ## Here if we are out of bounds: we calculate the amount of space tha
          ## we are overusing, divide that amongst our autoWidth columns 
          ## (columns which have to "width" key)
          
          ## Get the number of remaining columns' minimum width
          numColsRemaining = len(tableHeaders) - count
          widthDelta = abs(tableWidth - currentUsedWidth - headerWidth)
          remWidthDelta = abs(tableWidth - currentUsedWidth - headerWidth - (numColsRemaining * columnMinWidth))
          
          ## Divide the overflow (widthDelta) between all of our fixed columns
          ## (auto columns are all set to 20 at this point)
          #if not autoWidth:
            #numFixedColumns = len(fixedColumns) + 1
          #else:
            #numFixedColumns = len(fixedColumns)
          numFixedColumns = len(fixedColumns)
          if numFixedColumns > 0:            
            sharedDelta = remWidthDelta / numFixedColumns
          
            self.logger("New column width is out of bounds widthDelta:'%s' " % widthDelta
                      + "remWidthDelta:'%s' numFixedColumns:'%s'" % (remWidthDelta,numFixedColumns)
                      + " sharedDelta:'%s'" % sharedDelta, "debug")

            ## Go through them and subtract the sharedDelta
            for header in fixedColumns:
              header["actualWidth"] -= sharedDelta
          
          '''
          if not autoWidth:
            ## Apply the delta to our width only if we are not auto
            preferredRequestedWidth += headerWidth
            preferredUsedWidth += headerWidth - widthDelta             
            header["actualWidth"] = headerWidth - sharedDelta
          else:
            header["actualWidth"] = headerWidth
          '''
          currentUsedWidth += headerWidth - widthDelta
      
      ## At this point, we have calculated minimum widths for our columns
      ## If our total allocated space is less then our tableWidth, evently
      ## divide the remaining space amongst our autoWidth columns
      ## if no autoWidth columns exist, evently divide amongst our 
      ## fixed columns
      
      if currentUsedWidth < tableWidth:
        usedWidthDelta = tableWidth - currentUsedWidth
        
        if len(autoColumns) > 0:
          sharedDelta = usedWidthDelta / len(autoColumns)
          for header in autoColumns:
            header["actualWidth"] += sharedDelta
        else:
          sharedDelta = usedWidthDelta / len(tableHeaders)
          for header in tableHeaders:
            header["actualWidth"] += sharedDelta
          
               
      ## At this point, we have set the width for each column in the 
      ## "actualWidth" key.
    
      ## Iterate once again through all of our headers, append to new DOM     
      for header in tableHeaders:
        
        width = header["actualWidth"]
        preferredWidth = header["width"]
        ## Todo: progamatically determine based on byte count
        fittingWidth = header["width"]
        
        ## Debug
        #width = "155.66667175292969"
        #preferredWidth = "155.66667175292969"
        #fittingWidth = "62.689453125"
        
        gridColumnElement = xmlDoc.createElement("sf:grid-column")
        gridColumnElement.setAttribute("sf:fitting-width","%s" % fittingWidth)
        ## todo: figure out what 'sf:nc' is for...
        gridColumnElement.setAttribute("sf:nc", "3")
        gridColumnElement.setAttribute("sf:preferred-width",
                                        "%s" % preferredWidth)
        gridColumnElement.setAttribute("sf:width","%s" % width)
         
        self.logger("Finished processing header: '%s' '%s'" % (header["dataKey"],gridColumnElement.toxml()),"debug")
        ## Add the grid-column to our column entity
        columnsElement.appendChild(gridColumnElement)
      
    else:
      ## Here if no tableHeaders were provided: space will be equally
      ## provisioned for each of our columns, no header row will be
      ## drawn
      self.logger("Passing tables without tableHeaders is not currently supported!","error")
      return False
      avgWidth = tableWidth / numColumns
    
    gridElement.appendChild(columnsElement)
           
    ## <sf:vertical-gridline-styles sf:array-size="0"/>
    verticalGridElement = xmlDoc.createElement("sf:vertical-gridline-styles")
    verticalGridElement.setAttribute("sf:array-size","0")
    gridElement.appendChild(verticalGridElement)
    
    rowsElement = xmlDoc.createElement("sf:rows")
    rowsElement.setAttribute("sf:count","%s" % numRows)
      
    ## Run through our rows
    rowCount = numRows
    while rowCount > 0:
      ## Todo: programatically alter row height based on width, len column data
      rowHeight = minCellHeight
      gridRowElement = xmlDoc.createElement("sf:grid-row")
      gridRowElement.setAttribute("sf:fitting-height","25")
      gridRowElement.setAttribute("sf:height","%s" % rowHeight)
      ## Todo: figure out what these are for
      gridRowElement.setAttribute("sf:nc","2")
      gridRowElement.setAttribute("sf:ncoc","2")
      
      rowsElement.appendChild(gridRowElement)
      rowCount -= 1
      
    gridElement.appendChild(rowsElement)
    
    ## <sf:horizontal-gridline-styles sf:array-size="0"/>
    horizontalGridElement = xmlDoc.createElement("sf:horizontal-gridline-styles")
    horizontalGridElement.setAttribute("sf:array-size","0")
    gridElement.appendChild(horizontalGridElement)
    
    ## Build our data source
    dataSourceElement = xmlDoc.createElement("sf:datasource")
    
    headerCount = 0
    for header in tableHeaders:
      ## Calculate our dimensions after padding
      myHeaderHeight = minCellHeight - cellPaddingY
      myHeaderWidth = header["actualWidth"] - cellPaddingX
      ## debug
      ##myHeaderWidth = 45.337891
      
      headerCellElement = xmlDoc.createElement("sf:t")
      if headerCount == 0:
        headerCellElement.setAttribute("sf:ct","0")
      headerCellElement.setAttribute("sf:h","%s" % myHeaderHeight)
      headerCellElement.setAttribute("sf:s","%s" % standardHeaderStyleName)
      headerCellElement.setAttribute("sf:w","%s" % myHeaderWidth)
      headerCellDataElement = xmlDoc.createElement("sf:ct")
      if "displayName" in header and header["displayName"]:
        displayName = header["displayName"]
      else:
        displayName = header["dataKey"]
      headerCellDataElement.setAttribute("sfa:s","%s" % displayName)
      ##self.logger("Adding header cell: sf:ct='%s' sf:h='%s' sf:w='%s'" % (displayName,myHeaderHeight,myHeaderWidth),"debug")

      headerCellElement.appendChild(headerCellDataElement)      
      dataSourceElement.appendChild(headerCellElement)
      headerCount += 1
     
    ## Build our rows
    rowCount = 0
    for row in tableData:
      rowCount += 1
      self.logger("Adding row: %s" % rowCount,"debug")
      
      rowCellCount = 0
      

      myCellHeight = myHeaderHeight
      myCellWidth = myHeaderWidth
      ## debug
      ##myCellWidth = 60.689453
      
      for header in tableHeaders:
        rowCellElement = xmlDoc.createElement("sf:t")
        if "cellStyle" in header:
          cellStyle = header["cellStyle"]
        else:
          cellStyle = standardCellStyleName
          
        if rowCellCount == 0:
          ## Adjust our sf:ct for the first element of each row
          ## which seems to be 257 - numColumns, not sure significance of 257
          ctIndex = 257 - numColumns
          rowCellElement.setAttribute("sf:ct","%s" % ctIndex)
        rowCellElement.setAttribute("sf:h","%s" % myCellHeight)
        rowCellElement.setAttribute("sf:s","%s" % cellStyle)
        rowCellElement.setAttribute("sf:w","%s" % myCellWidth)
        rowCellDataElement = xmlDoc.createElement("sf:ct")       
        headerKey = header["dataKey"]
        if headerKey in row:
          cellData = row[headerKey]
        else:
          cellData = ""
          
        rowCellDataElement.setAttribute("sfa:s","%s" % cellData)
        #self.logger("Adding row:%s" % row,"debug")
        #self.logger("Adding row cell: '%s'" % cellData,"debug")

        rowCellElement.appendChild(rowCellDataElement)
        dataSourceElement.appendChild(rowCellElement)
        rowCellCount += 1

        
    
    gridElement.appendChild(dataSourceElement)
    tableModelElement.appendChild(gridElement)
    
    self.logger("Adding grid element.","debug")

    
    ## <sf:sort>
		##	<sf:sort-spec sf:sort-col="0" sf:sort-order="true"/>
		## </sf:sort>
    sortElement = xmlDoc.createElement("sf:sort")
    sortSpecElement = xmlDoc.createElement("sf:sort-spec")
    sortSpecElement.setAttribute("sf:sort-col","0")
    sortSpecElement.setAttribute("sf:sort-order","true")
    sortElement.appendChild(sortSpecElement)
    
    tableModelElement.appendChild(sortElement)
    
    ## <sf:filterset sf:enabled="false" sf:spec-count="1" sf:type="0">
    ##  <sf:filterspec sf:filter-col="65535" sf:key1="" sf:keyscale="0" sf:predicate="0"/>
    ## </sf:filterset>
    
    filterSetElement = xmlDoc.createElement("sf:filterset")
    filterSetElement.setAttribute("sf:enabled","false")
    filterSetElement.setAttribute("sf:spec-count","1")
    filterSetElement.setAttribute("sf:type","0")
    
    filterSpecElement = xmlDoc.createElement("sf:filterspec")
    filterSpecElement.setAttribute("sf:filter-col","65535")
    filterSpecElement.setAttribute("sf:key1","")
    filterSpecElement.setAttribute("sf:keyscale","0")
    filterSpecElement.setAttribute("sf:predicate","0")
    filterSetElement.appendChild(filterSpecElement)

    tableModelElement.appendChild(filterSetElement)
    
    ## <sf:cell-comment-mapping/>
    ## <sf:error_warning_mapping/>    
    tableModelElement.appendChild(xmlDoc.createElement("sf:cell-comment-mapping"))
    tableModelElement.appendChild(xmlDoc.createElement("sf:error_warning_mapping"))
  
    self.logger("Adding tableModelElement element to attachmentNode.","debug")

    tableMaster.appendChild(tableModelElement)
    
    return tableMaster
    
    
    
    
  def _summarizeTextStorageNode(self,DOM=None):
    """Creates a summary of text-storage node headers"""
    if not DOM:
      DOM = self.DOM
      
    if not DOM:
      self.logger("No DOM loaded to summarize!","error")
      return False
    
    ## Get our text-storage node
    textStorageNode = ""
    for element in DOM.childNodes:
      if element.nodeName == "sf:text-storage":
        textStorageNode = element
        break;
    for element in textStorageNode.childNodes:
      print element.nodeName
      for element1 in element.childNodes:
        print "\t%s" % element1.nodeName

    
    
      
      
      
      
      
    