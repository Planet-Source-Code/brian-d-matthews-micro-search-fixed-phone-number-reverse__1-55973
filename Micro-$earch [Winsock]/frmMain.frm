VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Micro-$earch - [Source]"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7110
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   7110
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Find a Person"
      Height          =   2895
      Left            =   3600
      TabIndex        =   13
      Top             =   120
      Width           =   3375
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   2160
         TabIndex        =   24
         Text            =   "GA"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Text            =   "30240"
         Top             =   1680
         Width           =   1935
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Begins With"
         Height          =   255
         Left            =   2160
         TabIndex        =   20
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Text            =   "Doe"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Begins With"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2160
         TabIndex        =   15
         Top             =   480
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Text            =   "John"
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "To be continued..."
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   2040
         Width           =   2895
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "State"
         Height          =   255
         Left            =   2160
         TabIndex        =   23
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "City, zip, or postal code"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1440
         Width           =   3135
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Last Name *"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "[* Indicates a required field ]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   2640
         Width           =   3135
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "First Name"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1935
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   120
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "www.whitepages.com"
      RemotePort      =   80
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Reverse Phone Number"
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1080
         MaxLength       =   7
         TabIndex        =   4
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         MaxLength       =   3
         TabIndex        =   1
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Command2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cancel"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1680
         TabIndex        =   8
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Command1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Process"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2520
         TabIndex        =   7
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Phone Number: [7-Digits]"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1080
         TabIndex        =   3
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Area Code"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Address"
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   3375
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   855
         IntegralHeight  =   0   'False
         ItemData        =   "frmMain.frx":0442
         Left            =   120
         List            =   "frmMain.frx":0444
         TabIndex        =   6
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2760
      Width           =   3375
   End
   Begin VB.Label SB1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Micro-$earch"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3120
      Width           =   3375
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   15
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   3375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' |---|-----------------------------------------|---|
'-[_________________________________________________|
'__[PlanetSourceCode.Com]___________________________)
'-[Micro Search]-: Created by: Brian Matthews [...]|)
'-[YCrack.Net]-: (C) Copyright 2004 YCrack : -[tm]-/|
'-[Credits:  http://YCrack.Net][http://YahPro.Org]-/|
'-[From: YahPro: keith_escalade [Analyze]:(Parsing)/|
'__[Microsoft-Murder_[Inc]_________________________/|
'__[Release_Date]:___[03/20/04]____________________/|
'____________________[Comments]____________________/|
'__Begin: [Declarations]___________________________/|
Private Declare Function InitCommonControls Lib "Comctl32.dll" () As Long
Dim AreaCode As String    '________________________/|
Dim PhoneNumber As String '________________________/|
Dim AllData As String     '________________________/|
'__End:   [Declarations]___________________________/|
'-[________________________________________________/|
'-[________________________________________________|)
' |---|-----------------------------------------|---|
Private Sub Form_Initialize()
Dim X As Long
X = InitCommonControls
End Sub
Private Sub Command1_Click()
If Text1.Text = "" Or Text2.Text = "" Then SB1.Caption = "Invalid Phone Number...": Exit Sub
If Len(Text1.Text) < 3 Or Len(Text2.Text) < 7 Then SB1.Caption = "Invalid Phone Number Length...": Exit Sub
Winsock1.Close
Winsock1.Connect
AreaCode = Text1.Text
PhoneNumber = Text2.Text
Command1.Enabled = False
Command2.Enabled = True
SB1.Caption = "Processing Phone Number: " & "(" & Text1 & ")" & " " & Text2 & "..."
Label4.Width = Label3.Width / 4
Label5.Caption = "25%"
End Sub
Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.BackStyle = 1
End Sub
Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.BackStyle = 0
End Sub
Private Sub Command2_Click()
Winsock1.Close
SB1.Caption = "Process Canceled..."
Command2.Enabled = False
Command1.Enabled = True
Label4.Width = 0
End Sub
Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command2.BackStyle = 1
End Sub
Private Sub Command2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command2.BackStyle = 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
Winsock1.Close: End
End Sub

Private Sub Text1_Change()
If IsNumeric(Text1.Text) = False Then SendKeys "{BackSpace}": Exit Sub
If Len(Text1.Text) = 3 Then Text2.SetFocus
End Sub
Private Sub Text2_Change()
If IsNumeric(Text2.Text) = False Then SendKeys "{BackSpace}": Exit Sub
End Sub

Private Sub Winsock1_Close()
On Error Resume Next
Label4.Width = Label3.Width
Label5.Caption = "100%"
Command1.Enabled = True
Command2.Enabled = False
If InStr(1, AllData, "Search Information:") Then
SB1.Caption = "Processing Complete...": ProcessAddress
Else
SB1.Caption = "No Results...": List1.Clear: Exit Sub
End If
Exit Sub
End Sub
Private Sub Winsock1_Connect()
On Error Resume Next
AllData = ""
Winsock1.SendData "GET http://www.whitepages.com/search/Reverse_Phone?npa=" & AreaCode & "&phone=" & PhoneNumber & vbCrLf & vbCrLf
Label4.Width = Label3.Width / 2
Label5.Caption = "50%"
End Sub
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim Data As String
Dim EXtra As Integer
Winsock1.GetData Data
AllData = AllData & Data
EXtra = Label3.Width / 100
Label4.Width = EXtra * 75
Label5.Caption = "75%"
End Sub
Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
Winsock1.Close
Command1.Enabled = True
Command2.Enabled = False
SB1.Caption = Description
Label4.Width = 0
End Sub
Private Sub ProcessAddress()
On Error Resume Next
Dim Needed As String
Dim FullName As String
List1.Clear
Needed = Get_Between(1, AllData, "<span class=""text"" style=""line-height:13pt;"">", "<br></span>")
FullName = Get_Between(1, AllData, "<img src=""/static/common/trans.gif"" width=""1"" height=""12"" border=""0""><a href=""javascript:moreInfo(1);""", "</td></tr>")
FullName = Replace(FullName, vbLf, "")
FullName = Replace(FullName, "</a>", "")
FullName = Replace(FullName, ">", "")
FullName = Replace(FullName, Chr(9), "")
FullName = Replace(FullName, ",", ", ")
FullName = Replace(FullName, "&amp;", "&")
List1.AddItem FullName
List1.AddItem Get_Item(1, Needed, "<br>")
List1.AddItem Get_Item(2, Needed, "<br>")
List1.AddItem Get_Item(3, Needed, "<br>")
End Sub
