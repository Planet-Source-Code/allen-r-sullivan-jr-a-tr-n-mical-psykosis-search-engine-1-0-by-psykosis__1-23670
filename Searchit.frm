VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1170
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4065
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Searchit.frx":0000
   ScaleHeight     =   1170
   ScaleWidth      =   4065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   1260
      Left            =   1680
      TabIndex        =   15
      Top             =   6480
      Width           =   1335
   End
   Begin VB.DirListBox Dir10 
      Height          =   315
      Left            =   840
      TabIndex        =   14
      Top             =   7440
      Width           =   855
   End
   Begin VB.DirListBox Dir9 
      Height          =   315
      Left            =   840
      TabIndex        =   13
      Top             =   7200
      Width           =   855
   End
   Begin VB.DirListBox Dir8 
      Height          =   315
      Left            =   840
      TabIndex        =   12
      Top             =   6960
      Width           =   855
   End
   Begin VB.DirListBox Dir7 
      Height          =   315
      Left            =   840
      TabIndex        =   11
      Top             =   6720
      Width           =   855
   End
   Begin VB.DirListBox Dir6 
      Height          =   315
      Left            =   840
      TabIndex        =   10
      Top             =   6480
      Width           =   855
   End
   Begin VB.DirListBox Dir5 
      Height          =   315
      Left            =   0
      TabIndex        =   9
      Top             =   7440
      Width           =   855
   End
   Begin VB.DirListBox Dir4 
      Height          =   315
      Left            =   0
      TabIndex        =   8
      Top             =   7200
      Width           =   855
   End
   Begin VB.DirListBox Dir3 
      Height          =   315
      Left            =   0
      TabIndex        =   7
      Top             =   6960
      Width           =   855
   End
   Begin VB.DirListBox Dir2 
      Height          =   315
      Left            =   0
      TabIndex        =   6
      Top             =   6720
      Width           =   855
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   0
      TabIndex        =   5
      Top             =   6480
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   3600
      Top             =   6000
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4680
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Height          =   255
      Left            =   3240
      TabIndex        =   1
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   5880
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Menu"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   495
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu frmUp 
         Caption         =   "Roll Up"
      End
      Begin VB.Menu frmDown 
         Caption         =   "Roll Down"
      End
      Begin VB.Menu brk2 
         Caption         =   "-"
      End
      Begin VB.Menu frmAbout 
         Caption         =   "About"
      End
      Begin VB.Menu brk1 
         Caption         =   "-"
      End
      Begin VB.Menu frmExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
List1.Clear
gimp$ = Replace(Text1, "*", "")
Timer1.Enabled = True
DoEvents
Dir1.Path = "C:\"
    DoEvents
    For a = 0 To Dir1.ListCount - 1
    Dir1.ListIndex = a
    Dir2.Path = Dir1.List(Dir1.ListIndex)
        For b = 0 To Dir2.ListCount - 1
        Dir2.ListIndex = b
        Dir3.Path = Dir2.List(Dir2.ListIndex)
            For c = 0 To Dir3.ListCount - 1
            Dir3.ListIndex = c
            Dir4.Path = Dir3.List(Dir3.ListIndex)
                For d = 0 To Dir4.ListCount - 1
                Dir4.ListIndex = d
                Dir5.Path = Dir4.List(Dir4.ListIndex)
                    For e = 0 To Dir5.ListCount - 1
                    Dir5.ListIndex = e
                    Dir6.Path = Dir5.List(Dir5.ListIndex)
                        For f = 0 To Dir6.ListCount - 1
                        Dir6.ListIndex = f
                        Dir7.Path = Dir6.List(Dir6.ListIndex)
                            For g = 0 To Dir7.ListCount - 1
                            Dir7.ListIndex = g
                            Dir8.Path = Dir7.List(Dir7.ListIndex)
                                For h = 0 To Dir8.ListCount - 1
                                Dir8.ListIndex = h
                                Dir9.Path = Dir8.List(Dir8.ListIndex)
                                    For i = 0 To Dir9.ListCount - 1
                                    Dir9.ListIndex = i
                                    Dir10.Path = Dir9.List(Dir9.ListIndex)
                                        For j = 0 To Dir10.ListCount - 1
                                        Dir10.ListIndex = j
                                        File1.Path = Dir10.List(Dir10.ListIndex)
                                            For k = 0 To File1.ListCount - 1
                                            File1.ListIndex = k
                                            If InStr(File1, gimp$) Then List1.AddItem (File1.Path + "\" + File1)
                                            DoEvents
                                            Next k
                                        DoEvents
                                        Next j
                                        File1.Path = Dir9.List(Dir9.ListIndex)
                                        For l = 0 To File1.ListCount - 1
                                        File1.ListIndex = l
                                        If InStr(File1, gimp$) Then List1.AddItem (File1.Path + "\" + File1)
                                        DoEvents
                                        Next l
                                    DoEvents
                                    Next i
                                    File1.Path = Dir8.List(Dir8.ListIndex)
                                    For M = 0 To File1.ListCount - 1
                                    File1.ListIndex = M
                                    If InStr(File1, gimp$) Then List1.AddItem (File1.Path + "\" + File1)
                                    DoEvents
                                    Next M
                                DoEvents
                                Next h
                                File1.Path = Dir7.List(Dir7.ListIndex)
                                For n = 0 To File1.ListCount - 1
                                File1.ListIndex = n
                                If InStr(File1, gimp$) Then List1.AddItem (File1.Path + "\" + File1)
                                DoEvents
                                Next n
                            DoEvents
                            Next g
                            File1.Path = Dir6.List(Dir6.ListIndex)
                            For o = 0 To File1.ListCount - 1
                            File1.ListIndex = o
                            If InStr(File1, gimp$) Then List1.AddItem (File1.Path + "\" + File1)
                            DoEvents
                            Next o
                        DoEvents
                        Next f
                        File1.Path = Dir5.List(Dir5.ListIndex)
                        For p = 0 To File1.ListCount - 1
                        File1.ListIndex = p
                        If InStr(File1, gimp$) Then List1.AddItem (File1.Path + "\" + File1)
                        Next p
                    DoEvents
                    Next e
                    File1.Path = Dir4.List(Dir4.ListIndex)
                    For r = 0 To File1.ListCount - 1
                    File1.ListIndex = r
                    If InStr(File1, gimp$) Then List1.AddItem (File1.Path + "\" + File1)
                    Next r
                DoEvents
                Next d
                File1.Path = Dir3.List(Dir3.ListIndex)
                For s = 0 To File1.ListCount - 1
                File1.ListIndex = s
                If InStr(File1, gimp$) Then List1.AddItem (File1.Path + "\" + File1)
                Next s
            DoEvents
            Next c
            File1.Path = Dir2.List(Dir2.ListIndex)
            For t = 0 To File1.ListCount - 1
            File1.ListIndex = t
            If InStr(File1, gimp$) Then List1.AddItem (File1.Path + "\" + File1)
            Next t
        DoEvents
        Next b
        File1.Path = Dir1.List(Dir1.ListIndex)
        For u = 0 To File1.ListCount - 1
        File1.ListIndex = u
        If InStr(File1, gimp$) Then List1.AddItem (File1.Path + "\" + File1)
        Next u
    DoEvents
    Next a
Timer1.Enabled = False
Label2 = "Finished " & List1.ListCount & " items found"
If Form1.Height < 6525 Then Label3.Caption = "Finished " & List1.ListCount & " items found"
End Sub

Private Sub Form_Load()
Form1.Height = 6525
Form1.Width = 4095
Form1.Visible = True
Call FormOnTop(Me)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
Call FormDrag(Me)
End Sub

Private Sub frmAbout_Click()
Call MsgBox("This was made because i've never seen it done before. I made an mp3 player based on this code and it works quite well. I am a newer programmer as i've only been programming for about 2 years now in visual basic, but i figure this might help you out just a little.", vbOKOnly + vbinformational + vbSystemModal, "psykosis search engine 1.0")
End Sub

Private Sub frmDown_Click()
If Form1.Height = 6525 Then Exit Sub
For i = Form1.Height To 6525
If Form1.Height >= 6520 Then Exit Sub
DoEvents
Form1.Height = Form1.Height + 5
DoEvents
Next i
End Sub

Private Sub frmExit_Click()
Me.Visible = False
DoEvents
Unload Me
End
End Sub

Private Sub frmUp_Click()
Do Until Form1.Height <= 1215
DoEvents
Form1.Height = Form1.Height - 5
DoEvents
Loop
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then Exit Sub
PopupMenu mnuFile, , Label1.Left, Label1.Top + Label1.Height
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
List1.Clear
gimp$ = Replace(Text1, "*", "")
Timer1.Enabled = True
DoEvents
Dir1.Path = "C:\"
    DoEvents
    For a = 0 To Dir1.ListCount - 1
    Dir1.ListIndex = a
    Dir2.Path = Dir1.List(Dir1.ListIndex)
        For b = 0 To Dir2.ListCount - 1
        Dir2.ListIndex = b
        Dir3.Path = Dir2.List(Dir2.ListIndex)
            For c = 0 To Dir3.ListCount - 1
            Dir3.ListIndex = c
            Dir4.Path = Dir3.List(Dir3.ListIndex)
                For d = 0 To Dir4.ListCount - 1
                Dir4.ListIndex = d
                Dir5.Path = Dir4.List(Dir4.ListIndex)
                    For e = 0 To Dir5.ListCount - 1
                    Dir5.ListIndex = e
                    Dir6.Path = Dir5.List(Dir5.ListIndex)
                        For f = 0 To Dir6.ListCount - 1
                        Dir6.ListIndex = f
                        Dir7.Path = Dir6.List(Dir6.ListIndex)
                            For g = 0 To Dir7.ListCount - 1
                            Dir7.ListIndex = g
                            Dir8.Path = Dir7.List(Dir7.ListIndex)
                                For h = 0 To Dir8.ListCount - 1
                                Dir8.ListIndex = h
                                Dir9.Path = Dir8.List(Dir8.ListIndex)
                                    For i = 0 To Dir9.ListCount - 1
                                    Dir9.ListIndex = i
                                    Dir10.Path = Dir9.List(Dir9.ListIndex)
                                        For j = 0 To Dir10.ListCount - 1
                                        Dir10.ListIndex = j
                                        File1.Path = Dir10.List(Dir10.ListIndex)
                                            For k = 0 To File1.ListCount - 1
                                            File1.ListIndex = k
                                            If InStr(File1, gimp$) Then List1.AddItem (File1.Path + "\" + File1)
                                            DoEvents
                                            Next k
                                        DoEvents
                                        Next j
                                        File1.Path = Dir9.List(Dir9.ListIndex)
                                        For l = 0 To File1.ListCount - 1
                                        File1.ListIndex = l
                                        If InStr(File1, gimp$) Then List1.AddItem (File1.Path + "\" + File1)
                                        DoEvents
                                        Next l
                                    DoEvents
                                    Next i
                                    File1.Path = Dir8.List(Dir8.ListIndex)
                                    For M = 0 To File1.ListCount - 1
                                    File1.ListIndex = M
                                    If InStr(File1, gimp$) Then List1.AddItem (File1.Path + "\" + File1)
                                    DoEvents
                                    Next M
                                DoEvents
                                Next h
                                File1.Path = Dir7.List(Dir7.ListIndex)
                                For n = 0 To File1.ListCount - 1
                                File1.ListIndex = n
                                If InStr(File1, gimp$) Then List1.AddItem (File1.Path + "\" + File1)
                                DoEvents
                                Next n
                            DoEvents
                            Next g
                            File1.Path = Dir6.List(Dir6.ListIndex)
                            For o = 0 To File1.ListCount - 1
                            File1.ListIndex = o
                            If InStr(File1, gimp$) Then List1.AddItem (File1.Path + "\" + File1)
                            DoEvents
                            Next o
                        DoEvents
                        Next f
                        File1.Path = Dir5.List(Dir5.ListIndex)
                        For p = 0 To File1.ListCount - 1
                        File1.ListIndex = p
                        If InStr(File1, gimp$) Then List1.AddItem (File1.Path + "\" + File1)
                        Next p
                    DoEvents
                    Next e
                    File1.Path = Dir4.List(Dir4.ListIndex)
                    For r = 0 To File1.ListCount - 1
                    File1.ListIndex = r
                    If InStr(File1, gimp$) Then List1.AddItem (File1.Path + "\" + File1)
                    Next r
                DoEvents
                Next d
                File1.Path = Dir3.List(Dir3.ListIndex)
                For s = 0 To File1.ListCount - 1
                File1.ListIndex = s
                If InStr(File1, gimp$) Then List1.AddItem (File1.Path + "\" + File1)
                Next s
            DoEvents
            Next c
            File1.Path = Dir2.List(Dir2.ListIndex)
            For t = 0 To File1.ListCount - 1
            File1.ListIndex = t
            If InStr(File1, gimp$) Then List1.AddItem (File1.Path + "\" + File1)
            Next t
        DoEvents
        Next b
        File1.Path = Dir1.List(Dir1.ListIndex)
        For u = 0 To File1.ListCount - 1
        File1.ListIndex = u
        If InStr(File1, gimp$) Then List1.AddItem (File1.Path + "\" + File1)
        Next u
    DoEvents
    Next a
Timer1.Enabled = False
Label2 = "Finished " & List1.ListCount & " items found"
If Form1.Height < 6525 Then Label3.Caption = "Finished " & List1.ListCount & " items found"
KeyAscii = 0
End If
End Sub

Private Sub Timer1_Timer()
Label2.Caption = File1.Path
End Sub
