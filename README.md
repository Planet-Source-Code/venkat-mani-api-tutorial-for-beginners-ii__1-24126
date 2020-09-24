<div align="center">

## API TUTORIAL for Beginners\-II


</div>

### Description

This is the 2nd in the series of articles and shows you the significance and importance of the sendmessage api and also showing you how the password text from a password field is captured using the sendmessage api
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Venkat Mani](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/venkat-mani.md)
**Level**          |Beginner
**User Rating**    |4.8 (29 globes from 6 users)
**Compatibility**  |VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/venkat-mani-api-tutorial-for-beginners-ii__1-24126/archive/master.zip)





### Source Code

<html>
<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>API TUTORIAL for Beginners</title>
</head>
<body bgcolor="#C0C0C0">
<p align="center"><font size="5">API TUTORIAL FOR BEGINNERS-II</font></p>
<p align="center"><font size="4" color="#000000">The SendMessage API</font></p>
<p align="center">&nbsp;</p>
<p align="left"><font color="#000000" size="3">The SendMessage Api is one of the
most powerful api functions . Before we&nbsp; take&nbsp; a look at it's uses and
syntax let me give you a brief overview of how the windows os works.</font></p>
<p align="left"><font color="#000000" size="3">The Windows Operating
System&nbsp; is a message based operating system .By saying message based means
that whenever the operating system (os) has to comunicate with applications
or&nbsp; two applications need to communicate/send data among themselves they do
so by sending messages to one another. For eg when an application is to be&nbsp;
terminated the os sends a WM_DESTROY message to that application, also when you
are adding an item to a listbox ,the application/os sends a LB_ADDSTRING message
to the listbox .&nbsp;&nbsp;</font></p>
<p align="left"><font color="#000000" size="3">While programming in VB the
sendmessage api is not of much use when u want to manipulate objects controls in
your own application.But say u wanted to change the title of some other
application or wanted to get the text from a textbox of another application or
want to terminate another application ,or set the text in a text box of another
application. The uses are endless if u want to&nbsp; play around with your
system.Also if you are planning to move over to win32 programming using&nbsp;
c++ you just cannot escape the sendmessage api.</font></p>
<p align="left"><font color="#000000" size="3">Let us look a the declaration of
the sendmessage api</font></p>
<p align="left">Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long<br>
</p>
<p align="left"><font color="#000000" size="3">&nbsp;The SendMessage api
function basically takes 4 parameters</font></p>
<ol>
 <li>
  <p align="left"><font color="#000000" size="3">hwnd-The handle of the window
  to which the message is being sent</font></li>
 <li>
  <p align="left"><font color="#000000" size="3">wMsg-The message that is
  being sent to the window.</font></li>
 <li>
  <p align="left"><font color="#000000" size="3">wParam-Parameter to be sent
  along with the message(depends on the message)</font></li>
 <li>
  <p align="left"><font color="#000000" size="3">lParam-Parameter to be sent
  along with the message(depends on the message)</font></li>
</ol>
<p align="left"><font color="#000000" size="3">Example1</font></p>
<p align="left"><font color="#000000" size="3">Let us see a practical
implementation of this api . Let us assume that we want&nbsp; to get the ***
masked text from a password textbox of a window!!! .We need to know a few things
before we can do this. The first thing we need to know is&nbsp; the handle
to&nbsp; the textbox window. One way of getting this is by using the
windowfrompoint api.Check my first tutorial on how to use this api and get the
window handle of the textbox.</font></p>
<p align="left"><font color="#000000" size="3">Once we have this handle we need
to send a WM_GETTEXTLENGTH message to the textbox .This message is essentially
sent&nbsp; to query the textbox and get the length of the text string in that
textbox.After we know the length of the string we have to send a WM_GETTEXT
message to the textbox and the textbox will return the text as the result .This
is how it is done</font></p>
<p align="left">Along with the declaration of the sendmessage api you need to
declare the 2 message&nbsp; constants that we are going to use</p>
<p align="left">Private Const WM_GETTEXT = &amp;HD<br>
<br>
Private Const WM_GETTEXTLENGTH = &amp;HE</p>
<p align="left">Put the following in any event of a control .In this example we
are putting it in a command click event</p>
<p align="left">Private Sub command1_click()&nbsp;</p>
<p align="left"><font color="#000000" size="3">Dim length As Long<br>
Dim result As Long<br>
Dim strtmp As String<br>
length = SendMessage(hwnd, WM_GETTEXTLENGTH, ByVal 0, ByVal 0) + 1<br>
strtmp = Space(length)<br>
result = SendMessage(hwnd, WM_GETTEXT, ByVal length, ByVal strtmp)<br>
</font></p>
<p align="left"><font color="#000000" size="3">End Sub</font></p>
<p align="left">here hwnd is the handle of the password textbox.</p>
<p align="left">Example 2</p>
<p align="left">In this example we will try to change the title of any
application ,in this case it will be a windows notepad application.</p>
<p align="left">As was the case previously we have to get the handle of the
notepad window .There are 2 ways to get this one is by using the windowfrompoint
api and the other is by using the findwindow api.The findwindow api returns the
handle of the window whose title has been specified in the function.</p>
<p align="left">After we get the handle of this window we do a sendmesaage
function</p>
<p align="left">dim result as long</p>
<p align="left">dim str1 as string</p>
<p align="left">str1=&quot;Venky&quot;</p>
<p align="left">result = SendMessage(hwnd, WM_SETTEXT, ByVal 0, ByVal str1)</p>
<p align="left">Using almost the similar techniques you can also put your own
text in the edit window of the notepad application.</p>
<p align="left">&nbsp;</p>
<p align="left">In this tutorial we have seen&nbsp; a few uses of the
sendmessage api.You can try out any of the numerous messages in the windows os
system on any applciation. Sendmessage is in other words a bridge for
communication between your application and another application</p>
<p align="left">Questions,comments send them to <a href="mailto:venky_dude@yahoo.com">venky_dude@yahoo.com</a>
</p>
<p align="left">&nbsp;</p>
<p align="center">&nbsp;</p>
<p align="left">&nbsp;</p>
</body>
</html>

