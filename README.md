# excel-corel-udl
Excel and CorelDRAW automation. VBA code for creating Corel layout from Excel data

Created for CorelDRAW X6 and Excel 2013.

Automates the task of drawing up files to be used for laser etching UDL's (Ultra Destructible Labels)

UDL's are used as serial number labels on products as they are extremely difficult to tamper with.

This Excel/Corel automation gives the option of creating the drawing from a column of values, either created or pasted into the document, or by providing a start and end serial number.

A CorelDRAW X6 file named "UDL_TEST.cdr" is included. This is required, or a CorelDRAW file with a Layer named "Main", a rectangle of size 20mm by 5mm name "Rect1", and artistic text named "Text1" with the string "X".

Sizes and file location can be changed in the PublicVariables module.

The colours red and green are the way that my laser is set up to recognize vectors for cutting.
My red uses less power than my green, so the red is to cut out the labels without going all the way through the material and the green is to cut the used piece away from the remaining material neatly.

If the text provided for the labels is larger than 18mm it will be compressed to fit and re-aligned to the centre of the label.

Changing the values in the PublicVariables module will allow you to customise the code to work with different sized labels quite easily.
