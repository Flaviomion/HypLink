# HypLink
HypLink is a VBscript which, when you have hyperlink on excel files, in the event of displacement of the cells, repositions them so as to maintain their functioning.
The script works on a copy of links, the TO link and the corresponding FR (from). In practice the first link connects us with a cell that can be in any sheet of the same file or in another file. The second link serves us as a return on the previous one.
To be able to reposition the damaged links by a displacement of cells the script needs metadata that is written on the cell comment.
This are examples of metadata:
  local   XL    HyFlink#XL#Pos=PS=xx#PC=jj#PR=rr#Punt=S=xx#C=jj#R=rr#HyElink
  Remote  TO    HyFlink#TO#Pos=PS=xx#PC=jj#PR=rr#cartella=C:\xxxxxx\yyyyy#file=LivelliAutonomia#S=xx#C=jj#R=rr#HyElink
	Remote  FR    HyFlink#FR#Pos=PS=xx#PC=jj#PR=rr#cartella=C:\xxxxxx\yyyyy#file=LivelliAutonomia#S=xx#C=jj#R=rr#HyElink
  
  
