from ScriptureObjects import ScriptureText
import re
import sys

Project = "BLKYI1"
OutputProject = "BLKYI1"
Books = "111111111111111111111111111111111111111111111111111111111111111111"
WordMedialPunctuation = ""
UpperCaseLetters = ""
LowerCaseLetters = ""
NoCaseLetters = ""
Separator = "|"
Diacritics = ""
DiacriticsFollow = "No"
FootnoteCallerSequence = ""
Encoding = "65001"
#Language = "ELisu"


def convertText(text):
    text2 = re.sub(r"\\Super\s+(\d[^ \\\s]*)\s*\\Super\*", r"\r\n\\v \1 ", text)
    text2 = re.sub(r"\\Super\s+(\d[^ \\\s]*) (.)\\Super\*", r"\r\n\\v \1 \2", text2)

    text2 = re.sub(ur"\uff0d", ur"-", text2)
    text2 = re.sub(ur"(\\c \d+\r\n\\p\S*)", ur"\1\r\n\\v 1 ", text2)
    text2 = re.sub(ur"\\Un", ur"\\pn", text2)
    text2 = re.sub(ur"(\\p\S*) \\Bd (.*)\\Bd\*\s*\r\n", ur"\1Bd \2\r\n", text2)
    text2 = re.sub(ur"\\[^\\s]+Super", ur"\\x1Super", text2)
    text2 = re.sub(ur"", ur"", text2)
     
    return text2

scr = ScriptureText(Project)     # open input project
if Project == OutputProject:
    scrOut = scr
else:
    scrOut = ScriptureText(OutputProject)	# open output project

flags = re.MULTILINE    # make ^ apply at the beginning of every line
# flags = flags | re.DOTALL   # make . match newline, comment this out if you don't want this

for reference, text in scr.chapters(Books):  # process all chapters
    sys.stderr.write(".")
    newText = convertText(text)
    scrOut.putText(reference, newText)
    if newText != text:
        sys.stderr.write(reference[:-2] + " changed\n")

# The books present might have changed so we need to update ssf file.
scrOut.save(OutputProject)
    
# Give the user a chance to see what has changed
#sys.stderr.write("\n\nPress <Enter> to end\n")
#sys.stdin.read(1)