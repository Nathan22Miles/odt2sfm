from elementtree.ElementTree import *
import re 
import codecs
import zipfile
import os
import glob

# This program converts open office document to use SFM codes.
# All the files to be converted must be in 'targetDirectory' below.
# They must have the extension .odt.
# The first 3 characters of the file name must be the usfm book id.
# The generated sfm text is placed in _GEN.sfm (for Genisis)
# RegExPal can be used to massage the conveted files into USFM
# after they have been imported into a Paratext project.
# This has only been tested with open office 2.0.

# All the generated paragraph styles will begin \p.
# The following strings will be added to the marker name based on the text in the paragraph:
#     It    Italic
#     Bd  Bold
#     SzN    Font size in points
#     StyleName      Open office style name if present)
#     InN     First line indent in 100ths of an inch
#     LftN   Left margin in 100ths of an inch
#     Cntr    Paragraph is centered
#
# Example: a bold centered paragraph .... \pBdCntr

# Character markers are constructed with the same scheme.
# InN, LftN, Cntr are never present for character markers.
# Character markers do not have an initial character p.
# Example: a bold italic string would be marked ... \BdIt ...\BdIt*
# Charater styles having default font size, non bold, non italic, with no style name are
# not marked in the sfm text.

targetDirectory = r"."

# Style attributes in this list will not be included in the sfm marker.
# Editing this list will make the sfm's simpler
defaultFontSizes = ["8pt", "9pt", "10pt", "10.5pt", "11pt", "12pt", "14pt", "16pt", "18pt", "39.5pt"]
defaultParentNames = ["Standard"]
defaultTextIndents = None    # None means ignore all indents
defaultMarginLefts = ["0in"]

# Name spaces for tags of interest
s = "{urn:oasis:names:tc:opendocument:xmlns:style:1.0}"    # open office styles
fo = "{urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0}"   # flow object formatting
ofc = "{urn:oasis:names:tc:opendocument:xmlns:office:1.0}"   # open office tags
txt = "{urn:oasis:names:tc:opendocument:xmlns:text:1.0}"    # open office textual data

# Return the tag of a node after stripping of its namespace.
def sTag(elem):
    return elem.tag.split("}")[1]

##############################################################
#  Analyze Open Office styles
##############################################################

# Return value of named attribute in this element or any subelement. "" if not present.
def getAttrValue(elem, name):
    if name in elem.attrib:
        return elem.attrib[name]
    
    for i in range(len(elem)):
        value = getAttrValue(elem[i], name)
        if value:
            return value
        
    return ""

# Info about an open document style
#     name
#     isParagraph - pargraph style (false means characters style)
#     isDropCap - contains drop cap chapter number
#     sfm - marker assigned to this style
class OdtStyle:
    # Extract attribute values from the style node.
    def __init__(self, elem):
        self.name = getAttrValue(elem, s+"name")
        
        family = getAttrValue(elem, s+"family")
        self.isParagraph = (family == "paragraph")
        if self.isParagraph:
            sfm = "p"
        else:
            sfm = ""
        
        self.isDropCap = getAttrValue(elem, s+"lines")
            
        parentName = getAttrValue(elem, s+"parent-style-name")
        if parentName and parentName not in defaultParentNames:
            parentName = parentName.replace("_20", "")
            sfm += parentName
            
        textPosition = getAttrValue(elem, s+"text-position");
        if textPosition and not textPosition.startswith("-"):
            sfm += "Super"
                        
        textAlign = getAttrValue(elem, fo+"text-align")
        if textAlign == "center":
            sfm += "Cntr"

        marginLeft = getAttrValue(elem, fo+"margin-left")
        if marginLeft and marginLeft not in defaultMarginLefts:
            margin = marginLeft[:-2]   # convert to nearest 100'th of an inch
            margin = int(float(margin) * 100)
            sfm += "Lft" + str(margin)

        textIndent = getAttrValue(elem, fo+"text-indent")
        if defaultTextIndents and textIndent and textIndent not in defaultTextIndents:
            indent = textIndent[:-2]   # convert to nearest 100'th of an inch
            indent = int(float(indent) * 100)
            sfm += "In" + str(indent)
        
        fontSize = getAttrValue(elem, fo+"font-size")
        if fontSize and fontSize not in defaultFontSizes:
            sfm += "Fs" + fontSize[:-2]

        fontStyle = getAttrValue(elem, fo+"font-style")
        if fontStyle == "italic":
            sfm += "It"

        fontWeight = getAttrValue(elem, fo+"font-weight")
        if fontWeight == "bold":
            sfm += "Bd"

        fontUnderline = getAttrValue(elem, s+"text-underline-style")
        if fontUnderline <> "":
            sfm += "Un"
            fontUnderlineType = getAttrValue(elem, s+"text-underline-type")
            if fontUnderlineType == "double":
                sfm += "Dbl"

            #print self.name.encode("utf-8") + " = " + sfm.encode("utf-8")
        self.sfm = sfm
        
# Generate a OdtStyle entry for each style in document.
# Store these by name in odtStyles dictionary.
def getStyles(root, odtStyles):
    for elem in root.getiterator():
        if sTag(elem) == "style":
            style = OdtStyle(elem)
            odtStyles[style.name] = style
            
    return odtStyles


##############################################################
#  Output XML text as SFM
##############################################################

# Write the text of this node and its subnotes
def outputNode(out, node, odtStyles, inFootnote):
    # Get the sfm marker for this node if any.
    # Find out if drop cap style.
    styleName = node.attrib.get(txt+"style-name", "")
    isDropCap = False
    if styleName:
        if styleName == "T1":
            isDropCap = False
        style = odtStyles.get(styleName, None)
        if style:
            sfm = style.sfm
            isDropCap = style.isDropCap
        else:
            # This should not happen unless there is a bug
            sfm = "UNKNOWN_" + styleName
    else:
        sfm = ""
    
    # If drop cap paragraph split off leading digits as the chapter number
    #! This will fail for ESG chapter A
    if isDropCap:
        if node.text <> None:
            cnum = node.text
        elif len(node) > 0 and node[0].tail <> None:
            cnum = node[0].tail
        else:
            cnum = ""
        parts = re.split(r"(\d+)", cnum, 1)
        if len(parts) == 3 and parts[0] == "":
            out.write("\r\n\\c " + parts[1])
            node.text = parts[2]
        else:
            print "Drop cap paragraph does not begin with a number"
    
    # Write initial sfm for paragraph, character style span, or note.
    tag = sTag(node)
    if tag == "p" or tag == "h":
        #! This only works for single paragraph footnotes
        if sfm and not inFootnote:
            out.write("\r\n\\" + sfm + " ")
    elif tag == "span":
        if sfm:
            out.write("\\" + sfm + " ")
    elif tag == "note":
        out.write("\\f + ")
        inFootnote = True
    elif tag == "tab":
        out.write("<tab>")

    # Write initial text of marker, if any
    if node.text:
        out.write(node.text)
    
    # Write text for all subnodes
    for i in range(len(node)):
        outputNode(out, node[i], odtStyles, inFootnote)
    
    # If character or note style marker, write end sfm tag
    if tag == "span" and sfm:
        out.write("\\" + sfm + "*")
    elif tag == "note":
        out.write("\\f*");
    
    # Write text following marker, if any
    if node.tail:
        out.write(node.tail)
        
def cleanupFootnote(match):
    text2 = match.group(0)
    
    text2 = re.sub(r"BdUn", "Un", text2)
    text2 = re.sub(ur"\\Bd[ *]", "", text2)
    
    text2 = re.sub(ur"\\f \+ \d+", ur"\\f + ", text2)
    text2 = re.sub(ur"\\f \+ \(\)", ur"\\f + ", text2)
    text2 = re.sub(ur"\\f \+ <tab>", ur"\\f + ", text2)
    return text2
        
def cleanupText(text):
    text2 = text
    
    #text2 = re.sub(ur"\\pIn50 (\d+\.)", ur"\\io2 \1", text2)
    #text2 = re.sub(ur"(\\p\w* )(?:\\Bd )?(\d+)\s*(?:\\Bd\*)?", ur"\r\n\r\n\\c \2\r\n\r\n\1", text2)
    
    text2 = re.sub(ur"\\Super\s+(\d[^\\]*)\s*\\Super\*", ur"\r\n\r\n\\v \1 ", text2)
    text2 = re.sub(ur"\\p \\Fs29 2\\Fs29\*", ur"\r\n\\c 2\r\n\\p ", text2)
    
    text2 = re.sub(ur"\\(Default_Paragraph_Font)?Super[ *]", ur"", text2)
    
    text2 = re.sub(ur"\\Bd\*(\s*)\\Bd ", ur"\1", text2)
    text2 = re.sub(ur"\\Un\*(\s*)\\Un ", ur"\1", text2)
    
    text2 = re.sub(ur"\\f .*?\\f\*", cleanupFootnote, text2)
        
    text2 = re.sub(ur"\(\s*\\f", ur"\\f", text2)
    text2 = re.sub(ur"\\f\*\s*\)", ur"\\f*", text2)
        
    text2 = re.sub(ur"", ur"", text2)
    text2 = re.sub(ur"", ur"", text2)
    text2 = re.sub(ur"", ur"", text2)
    text2 = re.sub(ur"", ur"", text2)
    text2 = re.sub(ur"", ur"", text2)
    text2 = re.sub(ur"", ur"", text2)
    
    return text2

# Output the text found under the office:text node
def outputText(root, odtStyles, bookId):
    node = root.find(".//" + ofc+ "text")
    if node == None:
        print "Cannot find text node"
        return
    
    out = codecs.open("_" + bookId + ".sfm", "w", "utf-8")
    
    out.write(r"\id " + bookId + "\r\n")    # write \id line    
    outputNode(out, root, odtStyles, False)    # write all nodes
    
    out.close()
    
    out = codecs.open("_" + bookId + ".sfm", "r", "utf-8")
    text = out.read()
    out.close()
    
    text2 = cleanupText(text)
    
    out = codecs.open("_" + bookId + ".sfm", "w", "utf-8")
    out.write(text2)
    out.close()
    


# Convert a single file.
# This generates _ + fileName[:3] + .sfm.
def convertFile(fileName):
    # Read stylesheet
    z = zipfile.ZipFile(fileName, "r")
    text = z.read("styles.xml")
    root = XML(text)
    odtStyles = {}
    getStyles(root, odtStyles)

    # Read locally modified styles from content.xml
    z = zipfile.ZipFile(fileName, "r")
    text = z.read("content.xml")
    root = XML(text)
    odtStyles = getStyles(root, odtStyles)
    
    # Output text using sfm styles
    outputText(root, odtStyles, fileName[:3])


# Convert all .odt files in the target directory
os.chdir(targetDirectory)
for fileName in glob.glob("*.odt"):
    print "Converting: " + fileName
    convertFile(fileName)
    