<div align="center">

## The joys of MsgBox's


</div>

### Description

This article will show you how to fully extend the basic VB MsgBox() function. Learn how to add icons, multiple lines of text, change the title, etc. MsgBox's are used in just about every application, learn how to use them to their maximum potential!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Dark Star X](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dark-star-x.md)
**Level**          |Beginner
**User Rating**    |3.9 (39 globes from 10 users)
**Compatibility**  |VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dark-star-x-the-joys-of-msgbox-s__1-23529/archive/master.zip)





### Source Code

Everyone uses MsgBox's, but most people just use it in the form of:<br><br>
MsgBox("Hello World")<br><br>
What most people don't realise is that the MsgBox function has a wide variety of functions.<br><br>
The proper syntax of the function is:<br><br>
MsgBox([Message], [options], [title], [Help File], [Context])<br><br>
For the purpose of not getting into too much advanced stuff here, we won't go into [context], we'll just leave it as vbNull, or not put anything in there at all.<br><br>
For [Message] You can put whatever you like surrounded in quotes. The basic string!<br>
For [options] you can put a combination of items using the Or keyword. [options] allows you to specify the buttons on the MsgBox (ok, cancel, retry, help, etc.), and/or the icon that appears on it. There are a bunch of already defined constants that VB allows you to use (seperate by the Or keyword if you want to use more than one, which I have already mentioned).<br><br>
vbCritical = The big red circle with an X in the middle of it<br>
vbExclamation = The yellow triangle with the exclamation point in it<br>
vbInformation = The white text bubble with the big i in it<br>
vbQuestion = The white text bubble with the big question mark in it<br><br>
<b>Note:</b> You cannot combine icons, if you try to, none of them will show up.<br><br>
You can also set which buttons show up, the names of the constants pretty much resemble what they do, for example:<br><br>
vbYesNo = Gives the user the option of hitting Yes or No<br>
vbOKOnly = Gives the user only the option to hit OK<br>
vbAbortRetryIgnore = Abort, Retry, and Ignore options are given.<br>
vbMsgBoxHelpButton = This will make a help button, and can be combined with vbOKOnly and vbCancelOnly.<br><br>
Try entering this into Visual Basic:<br>
Call MsgBox("Cannot Find File", vbCritical Or vbAbortRetryIgnore)<br>
Pretty neat eh?<br><br>
The last two options we should look at are:<br><br>
vbSystemModal, and vbApplicationModal<br>
vbApplicationModal prevents the user from preforming any work in the app until a button it pushed. vbSystemModal keeps the MsgBox on top of all windows until a button is pushed.<br><br>
The next setting is [title], this is pretty straight forward, as it sets the title for the MsgBox. If it is not set, the title will be the title of the application, NOT the form.<br><br>
[help file] is the path to the help file (in quotations), that will pop up when the user clicks the Help button if you include one. A help button will only be there if yoiu include a vbMsgBoxHelpButton in your options.<br><br>
Another great function of MsgBox's, are their ability to return values, for example, asking for confirmation, you could use some bit of code like:<br><br>
Dim X<br>
X = MsgBox("Are you sure?", vbInformation Or vbYesNo, "Confirmation")<br>
If X = vbYes Then Call MsgBox("Thanks", vbOKOnly)<br><br>Visual Basic also has a list of constants used to evaluate return values resulting from MsgBox's. They are pretty self explanitory, I'm sure you wouldn't even need to browse through a list to get most, if not all of them, some of them are as follows:<br><br>
vbYes = The 'Yes' button was pressed<br>
vbNo = The 'No' button was pressed<br>
vbCancel = The 'Cancel' button was pressed<br>
vbAbort = The 'About' button was pressed<br>
vbRetry = The 'Retry' button was pressed<br><br>
I'm sure you get the idea :)<br><br>
Last thing we're gonna look at, is making multiple lines of text. It's actually very very easy, by using the vbCrLf constant that VB has. When making your [Message] try throwing it in as you would a TextBox. For example:<br><br>
Call MsgBox("Hello" & vbCrLf & "World")<br><br><br>
Well, there's my input for today, you should all be experts on MsgBox's now, I hope I helped at least one person, I know it's a bit useless, but everyone uses MsgBox's, and it's good to know what you can do, just in case you get stuck.<br><br>
Please vote for me :)<br><br>
Regards,<br>
DarkStarX

