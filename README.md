# HypLink
HypLink is a VBscript which, when you have hyperlink on excel files, in the event of displacement of the cells, repositions them so as to maintain their functioning.
The script works on a copy of links, the TO link and the corresponding FR (from). In practice the first link connects us with a cell that can be in any sheet of the same file or in another file. The second link serves us as a return on the previous one.
To be able to reposition the damaged links by a displacement of cells the script needs metadata that is written on the cell comment.

This are examples of metadata:

  local
  HyFlink#XL#Pos=PS=xx#PC=jj#PR=rr#Punt=S=xx#C=jj#R=rr#HyElink
  
  Remote TO
  HyFlink#TO#Pos=PS=xx#PC=jj#PR=rr#cartella=C:\xxxxxx\yyyyy#file=LivelliAutonomia#S=xx#C=jj#R=rr#HyElink
  
  Remote FR
  HyFlink#FR#Pos=PS=xx#PC=jj#PR=rr#cartella=C:\xxxxxx\yyyyy#file=LivelliAutonomia#S=xx#C=jj#R=rr#HyElink
  
Use of the script

First: you need to create or have the excel files with the hyperlink as you need.
Second: create a file hyplink.ini with this contents:

	cartella=\\server\dir
	file=fileName
	rapporto=\\server\reportDir\report.html
	debug=no
	settaFont=si
	dimFont=11

Where
	cartella indicates the directory where the main excel file resides,
	file indicates the first part of the excel name (or the complete name),
	rapporto indicates where the report will be write,
	debug si means the debug is ON no means debug OFF,
	settaFonts si means that the script define a default dimension of the Comments font,
	dimFont is the dimension desired for the Comments Font.
	
Third: launch the script, it's look on .ini and read the excel files and create metadata to mantain the hyperlink found.

When you have performed these steps in case of modification of the files, you can re-execute the script and it will restore the hyperlinks that may have broken.

At last I create a Visual Basic version (hyplink_m1.vb) of HypLink to have an executable of the script.
