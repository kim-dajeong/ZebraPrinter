function for parsing zpl code and sending as bytes to printer on Windows.

static void ZebraPrinter.Printer.usbprint	(	string	variableNames,
                                            string	variables,
                                            string	templateFilename,
                                            string	usbDriverName )

string variableNames and variables should be in csv form with parameters ending in a circumflex (^). Ex) "one^,two^,three^"

ZebraDesigner 3 for Developers used to generate .prn print files with variables in zpl.

Example call in cmd:
cmdworkingdir>ZebraPrinter.exe "one^,two^,three^" "1^,2^,3^" "template.prn" "ZDesigner ZT411-600dpi ZPL"

Example zpl in template.prn:
CT~~CD,~CC^~CT~
^XA
~TA000
~JSN
^LT0
^MNW
^MTT
^PON
^PMN
^LH0,0
^JMA
^PR2,2
~SD15
^JUS
^LRN
^CI27
^PA0,1,1,0
^XZ
^XA
^MMT
^PW1800
^LL1200
^LS0
^FT176,372^A0N,163,162^FH\^CI28^FDone^FS^CI27
^FT176,577^A0N,163,162^FH\^CI28^FDtwo^FS^CI27
^FT176,801^A0N,163,162^FH\^CI28^FDthree^FS^CI27
^PQ1,,,Y
^XZ
