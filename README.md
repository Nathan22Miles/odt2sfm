Files to be converted must have the extension .odt.

The first 3 characters of the file name must be the usfm book id, e.g. "GEN".
The generated sfm text is placed in _GEN.sfm (for Genisis)

RegExPal can be used to massage the conveted files into USFM after they have been imported into a Paratext project.

All the generated paragraph styles will begin \p.
The following strings will be added to the marker name based on the text in the paragraph:

* It    Italic
* Bd  Bold
* SzN    Font size in points
* StyleName      Open office style name if present)
* InN     First line indent in 100ths of an inch
* LftN   Left margin in 100ths of an inch
* Cntr    Paragraph is centered

Example: a bold centered paragraph .... \pBdCntr

Character markers are constructed with the same scheme.
InN, LftN, Cntr are never present for character markers.
Character markers do not have an initial character p.

Example: a bold italic string would be marked ... \BdIt ...\BdIt*

Charater styles having default font size (see defaultFontSize in .py file), non bold, non italic, with no style name are
not marked in the sfm text.
