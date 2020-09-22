VERSION 5.00
Begin VB.Form HTMLgen 
   Caption         =   "HTMLgen"
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4695
   Icon            =   "HTMLgen.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Clear"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      ToolTipText     =   "Clear code area"
      Top             =   5760
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4440
      TabIndex        =   19
      Top             =   5760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      Height          =   375
      Left            =   480
      TabIndex        =   17
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Close"
      Height          =   495
      Left            =   3240
      TabIndex        =   9
      ToolTipText     =   "Get this of my screen"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Copy to Clipboard"
      Height          =   495
      Left            =   3240
      TabIndex        =   8
      ToolTipText     =   "Copy to clipboard"
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Open Notepad"
      Height          =   495
      Left            =   3240
      TabIndex        =   7
      ToolTipText     =   "Open Notepad"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate"
      Height          =   495
      Left            =   3240
      TabIndex        =   6
      ToolTipText     =   "Generate the code"
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox TxtSend 
      Height          =   2895
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   5
      ToolTipText     =   "Finished code"
      Top             =   2880
      Width           =   2775
   End
   Begin VB.TextBox TxtSize 
      Height          =   285
      Left            =   1680
      TabIndex        =   4
      Text            =   """+2"""
      ToolTipText     =   "Size of the font -6 to +6"
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox TxtColor 
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Text            =   """BLACK"""
      ToolTipText     =   "Color the font will be"
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox TxtImage 
      Height          =   285
      Left            =   720
      TabIndex        =   2
      Text            =   """Image here"""
      ToolTipText     =   "Image name and extension"
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox TxtBg 
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Text            =   """WHITE"""
      ToolTipText     =   "Color the page will be"
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox TxtTitle 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Page title"
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "If Image check box"
      Height          =   615
      Left            =   120
      TabIndex        =   18
      Top             =   960
      Width           =   735
   End
   Begin VB.Label lblfunc 
      BackStyle       =   0  'Transparent
      Caption         =   "Please make sure all fields are filled in before generating"
      Height          =   735
      Left            =   3120
      TabIndex        =   16
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Finished HTML Code"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   2640
      Width           =   2775
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Font Size ""-6 to +6"""
      Height          =   255
      Left            =   1680
      TabIndex        =   14
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Font Color "" """
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "BackGround Image "" """
      Height          =   255
      Left            =   720
      TabIndex        =   12
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "BackGround Color "" """
      Height          =   255
      Left            =   720
      TabIndex        =   11
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "HTMLgen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check2_Click()
Dim one, two, three, four
one = CStr("<BODY ") 'This decides if it will be a background image or color
two = CStr("BGCOLOR=") 'Demonstrates the if command and how to use a check box
three = CStr("BGIMAGE=")
four = CStr(">")
If Check2.Value = Checked Then
Text1 = one & three & TxtImage & four
ElseIf Check2.Value = Unchecked Then
Text1 = one & two & TxtBg & four
End If
End Sub

Private Sub Command1_Click()
Dim tag, tag1, tagrem, tag2, tag3, tag4, tag5, tag6, tag7, tag8, tag9, tag10
tag = CStr("<!--MADE WITH Steph's HTMLGenerator--!>")
tag1 = CStr("<HTML>")
tagrem = CStr("<!--PLEASE REMEMBER TO SAVE AS FILE EXTENTION .HTML--!>")
tag2 = CStr("<HEAD>")
tag3 = CStr("<TITLE>")
tag4 = CStr("</TITLE>") ' This is the command that generates all the code
tag5 = CStr("</HEAD>") 'Demonstrates how to use strings easily
tag6 = CStr("</BODY>") 'I didnt have to do it this way but figured it would be a good way to learn
tag7 = CStr("</HTML>")
tag8 = CStr("<FONT COLOR=")
tag9 = CStr(" SIZE=")
tag10 = CStr(">Enter Text Here</FONT>")
TxtSend.Text = tag1 & vbCrLf & tag & vbCrLf & tagrem & vbCrLf & tag2 & vbCrLf & tag3 & TxtTitle & tag4 & vbCrLf & tag5 & vbCrLf & Text1 & vbCrLf & tag6 & vbCrLf & tag8 & TxtColor & tag9 & TxtSize & tag10 & vbCrLf & tag6 & vbCrLf & tag7
End Sub

Private Sub Command2_Click()
MyAppID = Shell("C:\WINDOWS\NOTEPAD.EXE", 1)
AppActivate MyAppID ' Shell command to open notepad
End Sub

Private Sub Command3_Click()
Clipboard.SetText (TxtSend.Text) 'Copys text to clipboard
End Sub

Private Sub Command4_Click()
Unload Me 'Closes the program
End Sub

Private Sub Command5_Click()
TxtSend.Text = "" 'Clears the text field
End Sub

Private Sub Form_Unload(Cancel As Integer)
MsgBox "Thanx for the interest, Steph" 'This is how you make a message box open when the form unloads
End Sub
'Just a basic way to ease html page creation
'I was tired of writing all this stuff in notepad everytime
'And decided to share what I had done in hopes that I may help someone else learn
'How to do something to ease there workload
