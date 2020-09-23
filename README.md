<div align="center">

## The Daily Newbie \- Using the Dir\(\) Function


</div>

### Description

Explains the basics of using the Dir() command to get file and folder information.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Matthew Roberts](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/matthew-roberts.md)
**Level**          |Beginner
**User Rating**    |4.6 (23 globes from 5 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, VBA MS Access
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/matthew-roberts-the-daily-newbie-using-the-dir-function__1-23572/archive/master.zip)





### Source Code


<CENTER><FONT size="7">The Daily Newbie</FONT></CENTER>
<TABLE cellspacing="0" cellpadding="8" border="0">
	<TR>
		<TH colspan="2" align="middle" nowrap><BR><FONT face="Arial"><EM>To Start
   Things Off Right</EM><BR></FONT>
		</TH>
	</TR>
	<TR>
		<TD valign="top" width="30%"><B><FONT face="Arial">Today's
   Topic:</FONT> </B>
		</TD>
		<TD valign="top"><FONT face="Arial">Using the Dir
 Command</FONT>
		</TD>
	</TR>
	<TR>
		<TD valign="top"><B><FONT face="Arial">Name
   Derived From</FONT> </B>
		</TD>
		<TD valign="top"><FONT face="Arial">"Directory"</FONT>
		</TD>
	</TR>
	<TR>
		<TD valign="top"><B><FONT face="Arial">Used
   For:</FONT> </B>
		</TD>
		<TD valign="top"><FONT face="Arial">Getting Information about a
   particular folder or file.</FONT>
		</TD>
	</TR>
	<TR>
		<TD valign="top"><B><FONT face="Arial">VB Help
   File Description</FONT>  </B>
		</TD>
		<TD valign="top"><FONT face="Arial">Returns a String representing
   the name of a file, directory, or folder that matches a specified pattern
   or file attribute, or the volume label of a drive. Syntax </FONT>
		</TD>
	</TR>
	<TR>
		<TD valign="top"><B><FONT face="Arial">Plain
   English Description</FONT> </B>
		</TD>
		<TD valign="top">
<P><FONT face="Arial">Depending on the paramters set,
   returns:</FONT></P>
			<UL type="1">
				<LI><FONT face="Arial">The Name of a file in a folder that matches a
    patten (*.txt) </FONT>
				<LI><FONT face="Arial">The name of a sub folder within a folder. </FONT>
				<LI><FONT face="Arial">The name of a hard drive.</FONT></LI>
			</UL>
		</TD>
	</TR>
	<TR>
		<TD valign="top"><B><FONT face="Arial">Usage:</FONT></B>
		</TD>
		<TD valign="top">
<P><FONT face="Arial">To get a file name from a
   directory<FONT face="Courier">:&nbsp;&nbsp; strFile = Dir
   ("c:\MyFolder\*.txt")</FONT></FONT></P><FONT face="Courier">
<P><FONT face="Arial">To get a read-only file name from a directory<FONT face="Courier">:&nbsp;&nbsp; strFile = Dir ("c:\MyFolder\*.txt",
   vbReadOnly)</FONT></FONT></P>
<P><FONT face="Arial">To get a sub-directory from a directory<FONT face="Courier">:&nbsp;&nbsp; strFile = Dir ("c:\MyFolder\*",
   vbDirectory)</FONT></FONT></P>
<P><FONT face="Arial">To get the label of a drive<FONT face="Courier">:&nbsp;&nbsp; strFile = Dir ("d:",
   vbVolume)</FONT></FONT></P></FONT>
		</TD>
	</TR>
	<TR>
		<TD valign="top"><B><FONT face="Arial">Parameters:</FONT></B>
		</TD>
		<TD valign="top"><FONT face="Arial"></FONT>
			<UL>
				<LI>Path - The root directory to search from.
				<LI>Attribute - One of the following VB attribute
    values: <STRONG>vbNormal (default), vbReadOnly, vbHidden, VbSystem,
    vbVolume, vbDirectory</STRONG></LI>
			</UL>
<P>&nbsp;</P>
		</TD>
	</TR>
	<TR>
		<TD valign="top"><B><FONT face="Arial">Copy
   &amp; Paste Code:</FONT>  </B>
		</TD>
		<TD valign="top">
<P><FONT face="Arial">Today's copy and paste code lists
   all of the files in a directory to the debug window. For details on usage
   of the Dir ()command in this example, see the Notes below.</FONT></P>
		<PRE>  Dim strPathAndPattern As String<BR>&nbsp;&nbsp;&nbsp; Dim strFileName As String<BR>&nbsp;&nbsp;&nbsp;<BR>&nbsp;&nbsp;&nbsp; strPathAndPattern = InputBox("Enter a path and search pattern (ex: c:\windows\*.exe):")<BR>&nbsp;&nbsp;&nbsp;<BR>&nbsp;&nbsp;&nbsp;&nbsp;strFileName = Dir(strPathAndPattern)<BR>&nbsp;&nbsp;&nbsp;&nbsp;Debug.Print strFileName<BR>&nbsp;&nbsp;&nbsp;<BR>&nbsp;&nbsp;&nbsp;&nbsp;While strFileName &gt; ""<BR>&nbsp;&nbsp;&nbsp;&nbsp;   strFileName = Dir<BR>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Debug.Print strFileName
  Wend</CODE></PRE>
		</TD>
	</TR>
	<TR>
		<TD valign="top"><B><FONT face="Arial">Notes</FONT></B>
		</TD>
		<TD valign="top">
			<UL>
				<LI><FONT face="Arial">IMPORTANT: To find multiple
    files, you must do a "two step" proccess. In the example above, notice
    that the first time Dir() is called (before the While...Wend loop), the
    parameter <EM>strPathAndPatten </EM>is used. After that, the call is
    simply <EM>Dir</EM>. (strFileName = Dir). This is sort of confusing if
    you don't know what is going on. When I first used the Dir command, I
    kept getting the same file name over and over. This is because <U>using
    Dir() with a path parameter will always return the FIRST match</U>. To
    get subsequent matches, you simply call Dir(). It remembers the last
    pattern and passes the NEXT match. Really screwy.<BR></FONT>
				<LI><FONT face="Arial">Dir can be used to see if a file already
    exists:<BR><BR><FONT face="Courier">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; If Dir
    ("c:\MyFolder\log.txt") &gt; ""
    Then<BR>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    MsgBox "File Already
    Exists!"<BR>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; End
    If<BR></FONT></FONT><FONT face="Arial"><FONT face="Courier"><FONT face="Arial"><BR>This is because Dir() will return an empty string ("") if
    a match is not found, but will return the <U>file name</U> if a match
    <EM>is </EM>found.<BR></FONT></FONT></FONT>
				<LI><FONT face="Arial">Dir can also be used to see if a folder exists, but
    this is a little different. Because of the way Windows handles folder
    names, there is always a folder named "." and another named ".." . As
    odd as that seems, these names represent the current folder and the
    parent folder. So to find out if a certain folder exists, you can do
    this:<BR><BR><FONT face="Courier">If Dir ("c:\windows\system",
    vbDirectory) &gt; ".." Then<BR>&nbsp;&nbsp;&nbsp;&nbsp; MsgBox "Folder
    Already Exists!"<BR>End If</FONT><BR></FONT>
				<LI><FONT face="Arial">For a downloadable project using the Dir command,
    <A href="http://www.planetsourcecode.com/xq/ASP/txtCodeId.8369/lngWId.1/qx/vb/scripts/ShowCode.htm">Click Here</A></FONT></LI>
			</UL>
<P>&nbsp;</P>
		</TD>
	</TR>
	<TR>
		<TD valign="top"><FONT face="Arial"></FONT>
		</TD>
		<TD valign="top"><FONT face="Arial"></FONT>
		</TD>
	</TR>
	<TR>
		<TD colspan="2"><FONT face="Arial"></FONT>
		</TD>
	</TR>
</TABLE>

