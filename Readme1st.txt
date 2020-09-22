I decided to post my code in a little notepad type document for better functionality
OKOKOK i know my explanations are either too simple or too shitheaded to understand
Its hard to put out damn good code and then edit it to show everyone else tho, please understand this
If you do have any questions, feel free to email me @ neuropsykosis@aol.com
And please put the subject, "Planet source code-psykosis entry" or something like that so i don't delete yer shit
I ask that you do not leave something stupid as you may make the discarded list of gimps on the "bad sender" list
If ya vote for me good, if not, oh well, just give credit where credit is due
The bas file included is a portion of Dos32.bas by Chad J Cox and i am giving him credit here where it is duly deserved
i use his full bas all the time cuz i'm way too lazy to write my own.

BTW, don't luv me cuz i'm a gimp, n dun hate me cuz i'm leet

But enjoy the code and if ya use it, just put me somewhere in yer greetz or something NEXT TO ALL THE OTHER STOLEN CODE d;X just kidding LMFAO
Peace,
psykosis

---------------------------------------------------------------------------------------------------

Private Sub Command1_Click()
List1.Clear                                 <!~ clears the list
gimp$ = Replace(Text1, "*", "")             <!~ this way does not support the wildcard, so i removed it
Timer1.Enabled = True                       <!~ enables the timer so that you can see its working
DoEvents                                    <!~ lets everything kinda process
Dir1.Path = "C:\"                           <!~ paves the way for the first directory (and subsequencies later)
    DoEvents                                <!~ again that lil bugger d;]
    For a = 0 To Dir1.ListCount - 1         <!~ sets up the list for searching
    Dir1.ListIndex = a                      <!~ sets what is selected
    Dir2.Path = Dir1.List(Dir1.ListIndex)   <!~ makes the second dir list box a child of the original (ie. C:\ was 1, C:\windows would be 2)
        For b = 0 To Dir2.ListCount - 1     <!~ sets up second directory for lineup   
        Dir2.ListIndex = b                  <!~ sets what is to be fooled with (at the time)
        Dir3.Path = Dir2.List(Dir2.ListIndex) <!~ this is making dir3 a child of dir2 d;]
            For c = 0 To Dir3.ListCount - 1    <!~ you see where this is going?
            Dir3.ListIndex = c		       <!~ i'm going to skip to something new and let you ponder this code d;]

---------------------------------------------------------------------------------------------

File1.Path = Dir10.List(Dir10.ListIndex)     <!~ now this is different, since we've set the dir's (not all) we set the filebox 
      For k = 0 To File1.ListCount - 1         <!~ lineup, same as dirlistbox, but now it's filelistbox
      File1.ListIndex = k                      <!~ notice a pattern?
      If InStr(File1, gimp$) Then List1.AddItem (File1.Path + "\" + File1)  <!~ this actually looks at each file and asks if it is the right one d;]...if so it will add to list
      DoEvents                                 <!~ yep, lettin the program catch up w/ windows events
      Next k                                   <!~ goin onto the next one in the filebox
      DoEvents                                     <!~ letting windows catch up (again)
      Next j                                       <!~ this looks tricky, but actually this is from the dirlistbox previous to this, for every For there has to be a next so now we catch up w/ dirlistboxes
      File1.Path = Dir9.List(Dir9.ListIndex)       <!~ ahh, look familiar?
      For l = 0 To File1.ListCount - 1             <!~ still gettin repetitive?
      File1.ListIndex = l                          <!~ i know i know, getting old, moving on
      If InStr(File1, gimp$) Then List1.AddItem (File1.Path + "\" + File1)
      DoEvents
      Next l
                                    
--------------------------------------------------------------------------------------------

Timer1.Enabled = False                                   <!~ disables the timer
Label2 = "Finished " & List1.ListCount & " items found"  <!~ my own frill, just shows how many files it found
If Form1.Height < 6525 Then Label3.Caption = "Finished " & List1.ListCount & " items found" <!~ if the form is rolled up show it on other display
End Sub     <!~ ends the sub


---------------------------------------------------------------------------------------------

Private Sub Form_Load()
Form1.Height = 6525       <!~ sets the height for the form
Form1.Width = 4095        <!~ sets the width
Form1.Visible = True      <!~ makes it visible
Call FormOnTop(Me)        <!~ uses dos32's function formontop which sets the form on top of the screen
End Sub                   <!~ ends formload

-----------------------------------------------------------------------------------------------

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub        <!~ if not a leftclick we dont' want it
Call FormDrag(Me)                   <!~ will move the form (PHREEDOM FROM CAPTIONS!!)
End Sub                             <!~ BOOBYE SUB (can't ya tell im skitzo?)

-----------------------------------------------------------------------------------------------

Private Sub frmDown_Click()
If Form1.Height = 6525 Then Exit Sub     <!~ we dont' want the form rolling down too much
For i = Form1.Height To 6525             <!~ sets the basis for the rolldown
If Form1.Height >= 6520 Then Exit Sub    <!~ kind of "error correction for form rolling"
DoEvents                                 <!~ HA, BACK AGAIN!#!$@
Form1.Height = Form1.Height + 5          <!~ increases form height by 5
DoEvents                                 <!~ oops, its stalking you!!
Next i                                   <!~ goes back to the "for" statment
End Sub					 <!~ DIE SUB DIE@#!@ FIRE!!

-----------------------------------------------------------------------------------------------

Private Sub frmExit_Click()
Me.Visible = False                       <!~ so you don't see the form
DoEvents                                 <!~ lets windows catch up
Unload Me                                <!~ pulls program out of resident memory
End                                      <!~ ends the program, like the terminate command
End Sub                                  <!~ terminates the sub, any questions?

----------------------------------------------------------------------------------------------

Private Sub frmUp_Click()
if form1.height = 1215 then exit sub   <!~ more error correction
Do Until Form1.Height <= 1215     <!~ sets the rollup
DoEvents                           <!~ GIMPPPPPEYYYYYYYYYYYY
Form1.Height = Form1.Height - 5   <!~ sets the form only 5 notches higher
DoEvents                          <!~ i'm doevents, wanna play?
Loop                              <!~ goes back to the do statement
End Sub                           <!~ exits the sub

-------------------------------------------------------------------------------------------------
Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then Exit Sub       <!~ kind of error correction, i hate ppl goin "i right click for menu, its a bug" (don't ya wanna shoot em?)
PopupMenu mnuFile, , Label1.Left, Label1.Top + Label1.Height     <!~ sets the popup menu at the location i want it d;]
End Sub                      <!~ *cries* my sub died d;-(

--------------------------------------------------------------------------------------------------

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then    <!~ if the guy pressed the enter key then
List1.Clear                       <!~ removes anything done previous
gimp$ = Replace(Text1, "*", "")   <!~ that wildcard thing again
Timer1.Enabled = True             <!~ reenables the timer
DoEvents                          <!~ you know the routine, kinda the same i'll skip down
--------------------------------------------------------------------------------------------------

If Form1.Height < 6525 Then Label3.Caption = "Finished " & List1.ListCount & " items found"   <!~ same doodad d;]
KeyAscii = 0      <!~ this is kinda critical, if you do this you wont get that damn "DING!" sound when ya press the enter key, drove me mad for a week d;x
End If		<!~ hafta enter the end if, because you entered an if statment earlier (if keyascii = vbkeyreturn then) d;]
End Sub         <!~ *holds a memorial for his beloved sub*     LEAVE A FLOWER DAMMIT!!
----------------------------------------------------------------------------------------------------

Private Sub Timer1_Timer()
Label2.Caption = File1.Path    <!~ remember when you enabled the timer? well this is what the timer is doing, its displaying every hell, 500th of a second what dir its looking in d;]
End Sub
