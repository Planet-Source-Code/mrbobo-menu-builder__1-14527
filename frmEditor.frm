VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bobo Menu Builder"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   5415
   Icon            =   "frmEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox rtftext 
      Height          =   255
      Left            =   1560
      TabIndex        =   33
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393217
      TextRTF         =   $"frmEditor.frx":08CA
   End
   Begin MSComctlLib.ListView LV 
      Height          =   255
      Left            =   3840
      TabIndex        =   21
      Top             =   480
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000011&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   5415
      TabIndex        =   28
      Top             =   0
      Width           =   5415
      Begin VB.TextBox Text4 
         BackColor       =   &H8000000B&
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   1320
         TabIndex        =   30
         Top             =   120
         Width           =   3735
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H8000000B&
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   1320
         TabIndex        =   29
         Top             =   495
         Width           =   3735
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000010&
         Caption         =   "Form Name :"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   165
         Width           =   975
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000010&
         Caption         =   "Path :"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   525
         Width           =   855
      End
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "frmEditor.frx":0993
      Left            =   3960
      List            =   "frmEditor.frx":09A3
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Height          =   315
      Left            =   1320
      TabIndex        =   24
      Text            =   "0"
      Top             =   2280
      Width           =   975
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Widow List"
      Height          =   255
      Left            =   3840
      TabIndex        =   23
      Top             =   2700
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   0
      ScaleHeight     =   2655
      ScaleWidth      =   135
      TabIndex        =   22
      Top             =   3480
      Width           =   135
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmEditor.frx":09D2
      Left            =   3360
      List            =   "frmEditor.frx":0AC6
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1880
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   840
      TabIndex        =   2
      Top             =   1880
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete"
      Height          =   330
      Left            =   4200
      TabIndex        =   13
      Top             =   3180
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Insert"
      Height          =   330
      Left            =   3000
      TabIndex        =   12
      Top             =   3180
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Next"
      Height          =   330
      Left            =   1800
      TabIndex        =   11
      Top             =   3180
      Width           =   1095
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Visible"
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   2700
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Enabled"
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   2700
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Checked"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2700
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   4200
      TabIndex        =   16
      Top             =   1480
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   330
      Left            =   4200
      TabIndex        =   15
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Top             =   1480
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   840
      TabIndex        =   0
      Top             =   1080
      Width           =   3135
   End
   Begin VB.CommandButton cmdpos 
      Height          =   330
      Index           =   3
      Left            =   1200
      Picture         =   "frmEditor.frx":0DA5
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3180
      Width           =   300
   End
   Begin VB.CommandButton cmdpos 
      Height          =   330
      Index           =   2
      Left            =   840
      Picture         =   "frmEditor.frx":10D3
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3180
      Width           =   300
   End
   Begin VB.CommandButton cmdpos 
      Height          =   330
      Index           =   1
      Left            =   480
      Picture         =   "frmEditor.frx":1401
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3180
      Width           =   300
   End
   Begin VB.CommandButton cmdpos 
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   120
      Picture         =   "frmEditor.frx":172F
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3180
      Width           =   300
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4800
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   5
   End
   Begin VB.Label Label7 
      Caption         =   "NegotiatePosition :"
      Height          =   255
      Left            =   2520
      TabIndex        =   27
      Top             =   2340
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "HelpcontextID :"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   2340
      Width           =   1335
   End
   Begin MSForms.ListBox List1 
      Height          =   2415
      Left            =   0
      TabIndex        =   14
      Top             =   3600
      Width           =   5295
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "9340;4260"
      ColumnCount     =   2
      cColumnInfo     =   1
      MatchEntry      =   0
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "6350"
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   5280
      Y1              =   3045
      Y2              =   3045
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5280
      Y1              =   3030
      Y2              =   3030
   End
   Begin VB.Label Label4 
      Caption         =   "Shortcut :"
      Height          =   255
      Left            =   2520
      TabIndex        =   20
      Top             =   1925
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Index :"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   1925
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Name :"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   1520
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Caption :"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   1110
      Width           =   735
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "New Menu"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open Form"
      End
      Begin VB.Menu mnuFileOpenTemplate 
         Caption         =   "Open Template"
      End
      Begin VB.Menu mnuFileSpace 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSaveFormAs 
         Caption         =   "Save Form As"
      End
      Begin VB.Menu mnuFileSaveMenu 
         Caption         =   "Save Menu As New Form"
      End
      Begin VB.Menu mnuFileSaveAsTemplate 
         Caption         =   "Save Menu As Template"
      End
      Begin VB.Menu mnuFileSpace1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuEditClear 
         Caption         =   "Clear Menu"
      End
      Begin VB.Menu mnuEditTemplate 
         Caption         =   "Replace with Template"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
   End
End
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright Bobo Enterprises 2001
'This is a beta version of a tool which forms part of a commercial
'release VB6 addin. This version is made as a stand-alone exe for
'testing. Some of the code is a bit messy and inefficient.
'Most of the code is self explanatory or is simple 'House keeping'
'and I haven't bothered to comment on it.

'Recommend you test it first on copies of forms to
'get the hang of how it works

'***ADVANTAGES OVER STANDARD MENU EDITOR***
'No limit on size or nested submenus
'Allows easy moving of menu structures between forms
'Lets you save oft used menus for re-use


'***DISADVANTAGES***
'This beta works outside the IDE

'I've included the couple of images used so just
'compile the EXE and you should have a useful tool.
'Please send any comments or report bugs to
'gtkerr@bigpond.com




Public existing As Boolean          'it's an existing form we're editing
Public ExistingPath As String       'and this is where its' at
Dim ic As ListItem
Dim InvalidMenu As Boolean          'they cocked up, submenu in the wrong place or summit
Dim BeforeTxt As String             'the text in a form before the menu structure
Dim AfterTxt As String              'the text in a form after the menu structure
Dim curtext As String               'the menu structure
Dim textfound As Long
Dim pos As Long
Private Sub Check1_Click()
LV.SelectedItem.SubItems(1) = Check1.Value
End Sub
Private Sub Check2_Click()
LV.SelectedItem.SubItems(2) = Check2.Value
End Sub
Private Sub Check3_Click()
LV.SelectedItem.SubItems(3) = Check3.Value
End Sub
Private Sub Check4_Click()
LV.SelectedItem.SubItems(6) = Check4.Value
End Sub
Private Sub cmdCancel_Click()
Unload Me
End Sub
Private Sub cmdOK_Click()
'In a normal app this button would be the "Save" menuitem
'But to keep it like VB6s' menu editor we've used the "OK" button
On Error GoTo woops
Dim temp As String, sfile As String, myMenu As String
Dim DialogType As Integer
Dim DialogTitle As String
Dim DialogMsg As String
Dim Response As Integer
If Label5 = "Template" Then
    Screen.MousePointer = 11
    myMenu = GetMyMenu
    If InvalidMenu Then
        InvalidMenu = False
        Exit Sub
    End If
    Screen.MousePointer = 0
    FileSave myMenu, Text6.Text
    Exit Sub
End If
If List1.List(List1.ListCount - 1) = "" And LV.ListItems(LV.ListItems.Count).Text = "" Then
    List1.RemoveItem List1.ListCount - 1
    LV.ListItems.Remove LV.ListItems.Count
End If
If List1.ListCount = 0 Then
    myMenu = ""
Else
    Screen.MousePointer = 11
    myMenu = GetMyMenu
    Screen.MousePointer = 0
End If
If InvalidMenu Then
    InvalidMenu = False
    Exit Sub
End If
If existing = True Then
    'Better remind them this is editing a form
    'and cant be undone
    DialogType = vbYesNoCancel
    DialogTitle = "Bobo Enterprises"
    DialogMsg = "This will overwrite an existing form. Do you wish to save as a copy instead ?"
    Response = MsgBox(DialogMsg, DialogType, DialogTitle)
    Select Case Response
        Case vbYes
            'Whooh ! Lets just save a copy to be safe
            With CommonDialog1
                .FileName = Text4.Text + ".frm"
                .DialogTitle = "Save Form"
                .CancelError = True
                .Filter = "VB 6 Forms |*.frm"
                .ShowSave
                If Len(.FileName) = 0 Then Exit Sub
                sfile = .FileName
            End With
        Case vbNo
            'Damn the torpedoes full speed ahead
            sfile = ExistingPath
        Case vbCancel
            'Panic
            Exit Sub
    End Select
Else
    With CommonDialog1
        .FileName = Text4.Text + ".frm"
        .DialogTitle = "Save Form"
        .CancelError = True
        .Filter = "VB 6 Forms |*.frm"
        .ShowSave
        If Len(.FileName) = 0 Then Exit Sub
        sfile = .FileName
    End With
End If
If Not existing Then
'The user wants a new form so lets create one
temp = "VERSION 5.00" + vbCrLf + "Begin VB.Form " + Text4.Text + vbCrLf + "   Caption         =   " + Chr(34) + Text4.Text + Chr(34) + vbCrLf _
+ "   ClientHeight    =   3195" + vbCrLf + "   ClientLeft      =   60" + vbCrLf + "   ClientTop       =   345" + vbCrLf _
+ "   ClientWidth     =   4680" + vbCrLf + "   LinkTopic       =   " + Chr(34) + "Form1" + Chr(34) + vbCrLf + "   ScaleHeight     =   3195" + vbCrLf _
+ "   ScaleWidth      =   4680" + vbCrLf + "   StartUpPosition =   3" + vbCrLf + myMenu + vbCrLf + "End" + vbCrLf _
+ "Attribute VB_Name = " + Chr(34) + Text4.Text + Chr(34) + vbCrLf + "Attribute VB_GlobalNameSpace = False" + vbCrLf + "Attribute VB_Creatable = False" + vbCrLf _
+ "Attribute VB_PredeclaredId = True" + vbCrLf + "Attribute VB_Exposed = False" + vbCrLf
Else
temp = BeforeTxt + myMenu + vbCrLf + "End" + vbCrLf + AfterTxt
End If
FileSave temp, sfile
existing = True
Text6.Text = sfile
ExistingPath = sfile
woops:
End Sub
Private Sub cmdpos_Click(Index As Integer)
Dim nItem As Integer
Select Case Index
Case 0 'left
    If Left(List1.List(List1.ListIndex), 4) = "····" Then
        List1.List(List1.ListIndex) = Right(List1.List(List1.ListIndex), Len(List1.List(List1.ListIndex)) - 4)
    End If
Case 1 'right
    List1.List(List1.ListIndex) = "····" + List1.List(List1.ListIndex)
Case 2 'up
    If List1.ListIndex < 1 Then Exit Sub
    nItem = List1.ListIndex
    If nItem = 0 Then Exit Sub
    Set ic = LV.ListItems.Add(nItem, , Text2.Text)
    ic.SubItems(1) = Check1.Value
    ic.SubItems(2) = Check2.Value
    ic.SubItems(3) = Check3.Value
    ic.SubItems(4) = Text3.Text
    ic.SubItems(5) = Combo1.ListIndex
    ic.SubItems(6) = Check4.Value
    ic.SubItems(7) = Combo2.ListIndex
    ic.SubItems(8) = Text5.Text
    LV.ListItems.Remove nItem + 2
    List1.AddItem List1.Text, nItem - 1
    List1.RemoveItem nItem + 1
    List1.Selected(nItem - 1) = True
Case 3 'down
    If List1.ListIndex < List1.ListCount - 1 Then
        nItem = List1.ListIndex
        If nItem = List1.ListCount - 1 Then Exit Sub
        Set ic = LV.ListItems.Add(nItem + 3, , Text2.Text)
        ic.SubItems(1) = Check1.Value
        ic.SubItems(2) = Check2.Value
        ic.SubItems(3) = Check3.Value
        ic.SubItems(4) = Text3.Text
        ic.SubItems(5) = Combo1.ListIndex
        ic.SubItems(6) = Check4.Value
        ic.SubItems(7) = Combo2.ListIndex
        ic.SubItems(8) = Text5.Text
        LV.ListItems.Remove nItem + 1
        List1.AddItem List1.Text, nItem + 2
        List1.RemoveItem nItem
        List1.Selected(nItem + 1) = True
    Else
        If List1.List(List1.ListCount - 1) <> "" Then
            List1.AddItem ""
            Text2.Text = ""
            Set ic = LV.ListItems.Add(, , Text2.Text)
            ic.SubItems(1) = Check1.Value
            ic.SubItems(2) = Check2.Value
            ic.SubItems(3) = Check3.Value
            ic.SubItems(4) = Text3.Text
            ic.SubItems(5) = Combo1.ListIndex
            ic.SubItems(6) = Check4.Value
            ic.SubItems(7) = Combo2.ListIndex
            ic.SubItems(8) = Text5.Text
            List1.ListIndex = List1.ListIndex + 1
        End If
    End If
End Select
'update caption and name
Text1.Text = Mid$(List1.List(List1.ListIndex), InStrRev(List1.List(List1.ListIndex), "·") + 1)
Text2.Text = ic.Text
End Sub
Private Sub Combo1_Click()
LV.SelectedItem.SubItems(5) = Combo1.ListIndex
If Combo1.ListIndex > 0 Then
    List1.Column(1, List1.ListIndex) = Combo1.Text
Else
    List1.Column(1, List1.ListIndex) = ""
End If
End Sub
Private Sub Combo2_Click()
LV.SelectedItem.SubItems(7) = Combo2.ListIndex
End Sub
Private Sub Command1_Click() 'next
Dim emp As String
If List1.ListIndex < List1.ListCount - 1 Then
    List1.ListIndex = List1.ListIndex + 1
Else
    If List1.List(List1.ListCount - 1) <> "" Then
        emp = Mid$(List1.List(List1.ListCount - 1), 1, InStrRev(List1.List(List1.ListCount - 1), "·"))
        List1.AddItem emp
        Combo1.ListIndex = 0
        Check1.Value = 0
        Check2.Value = 1
        Check3.Value = 1
        Check4.Value = 0
        Combo2.ListIndex = 0
        Text5.Text = "0"
        Set ic = LV.ListItems.Add(, , "")
        ic.SubItems(1) = Check1.Value
        ic.SubItems(2) = Check2.Value
        ic.SubItems(3) = Check3.Value
        ic.SubItems(4) = Text3.Text
        ic.SubItems(5) = Combo1.ListIndex
        ic.SubItems(6) = Check4.Value
        ic.SubItems(7) = Combo2.ListIndex
        ic.SubItems(8) = Text5.Text
        Dim bg As Integer
        bg = LV.ListItems.Count
        List1.ListIndex = List1.ListIndex + 1
    End If
End If
Text1.Text = Mid$(List1.List(List1.ListIndex), InStrRev(List1.List(List1.ListIndex), "·") + 1)
Text2.Text = ic.Text
End Sub
Private Sub Command2_Click() 'insert
Dim emp As String
emp = Mid$(List1.List(List1.ListIndex), 1, InStrRev(List1.List(List1.ListIndex), "·"))
List1.AddItem emp, List1.ListIndex
Combo1.ListIndex = 0
Check1.Value = 0
Check2.Value = 1
Check3.Value = 1
Check4.Value = 0
Combo2.ListIndex = 0
Text5.Text = "0"
Set ic = LV.ListItems.Add(, , "")
ic.SubItems(1) = Check1.Value
ic.SubItems(2) = Check2.Value
ic.SubItems(3) = Check3.Value
ic.SubItems(4) = Text3.Text
ic.SubItems(5) = Combo1.ListIndex
ic.SubItems(6) = Check4.Value
ic.SubItems(7) = Combo2.ListIndex
ic.SubItems(8) = Text5.Text
List1.ListIndex = List1.ListIndex - 1
Text1.Text = Mid$(List1.List(List1.ListIndex), InStrRev(List1.List(List1.ListIndex), "·") + 1)
Text2.Text = ic.Text
End Sub
Private Sub Command3_Click() 'delete
If List1.ListCount > 1 Then
    If List1.ListIndex > 0 Then
        List1.ListIndex = List1.ListIndex - 1
        List1.RemoveItem List1.ListIndex + 1
        LV.ListItems.Remove List1.ListIndex + 2
    Else
        List1.ListIndex = List1.ListIndex + 1
        List1.RemoveItem List1.ListIndex - 1
        LV.ListItems.Remove List1.ListIndex
    End If
Else
    List1.List(0) = ""
    LV.ListItems.Clear
    Combo1.ListIndex = 0
    Check1.Value = 0
    Check2.Value = 1
    Check3.Value = 1
    Check4.Value = 0
    Combo2.ListIndex = 0
    Text5.Text = "0"
    Set ic = LV.ListItems.Add(, , "")
    ic.SubItems(1) = Check1.Value
    ic.SubItems(2) = Check2.Value
    ic.SubItems(3) = Check3.Value
    ic.SubItems(4) = Text3.Text
    ic.SubItems(5) = Combo1.ListIndex
    ic.SubItems(6) = Check4.Value
    ic.SubItems(7) = Combo2.ListIndex
    ic.SubItems(8) = Text5.Text
End If
Text1.Text = Mid$(List1.List(List1.ListIndex), InStrRev(List1.List(List1.ListIndex), "·") + 1)
Text2.Text = ic.Text
End Sub
Private Sub Form_Load()
Dim mycommand As String
Dim temp As String
'Associates itself to its own filetype .bmu
'These are the template files to hold menu structures
'When clicked on in Explorer they open in this app
mycommand = Command()
If mycommand = "" Then 'not shelled so set defaults
    Text4.Text = "Form1"
    List1.AddItem ""
    Check1.Value = 0
    Check2.Value = 1
    Check3.Value = 1
    Check4.Value = 0
    Text5.Text = "0"
    Set ic = LV.ListItems.Add(, , Text2.Text)
    ic.SubItems(1) = Check1.Value
    ic.SubItems(2) = Check2.Value
    ic.SubItems(3) = Check3.Value
    ic.SubItems(4) = Text3.Text
    ic.SubItems(5) = 0
    ic.SubItems(6) = Check4.Value
    ic.SubItems(7) = 0
    ic.SubItems(8) = Text5.Text
    ic.Selected = True
    List1.ListIndex = 0
    Combo1.ListIndex = 0
    Combo2.ListIndex = 0
Else 'shelled so open the file and read the menu structure
    Text4.Text = Mid$(mycommand, InStrRev(mycommand, "\") + 1)
    Text6.Text = mycommand
    Label5 = "Template"
    rtftext.LoadFile mycommand  'using a Richtextbox to open files
                                'this avoids some errors
    curtext = rtftext.Text
    ParseMenu
End If
'make sure we're still associated to our filetype
Associate App.Path + "\BoboMenuBuilder.exe", ".bmu"
End Sub
Private Sub List1_Click()
Text1.Text = Mid$(List1.List(List1.ListIndex), InStrRev(List1.List(List1.ListIndex), "·") + 1)
LV.ListItems(List1.ListIndex + 1).Selected = True
Text2.Text = LV.SelectedItem.Text
Check1.Value = LV.SelectedItem.SubItems(1)
Check2.Value = LV.SelectedItem.SubItems(2)
Check3.Value = LV.SelectedItem.SubItems(3)
Text3.Text = LV.SelectedItem.SubItems(4)
Check4.Value = LV.SelectedItem.SubItems(6)
Text5.Text = LV.SelectedItem.SubItems(8)
Combo1.ListIndex = LV.SelectedItem.SubItems(5)
Combo2.ListIndex = LV.SelectedItem.SubItems(7)
End Sub
Private Sub mnuEditClear_Click()
List1.Clear
LV.ListItems.Clear
List1.AddItem ""
Set ic = LV.ListItems.Add(, , "")
ic.SubItems(1) = 0
ic.SubItems(2) = 1
ic.SubItems(3) = 1
ic.SubItems(4) = ""
ic.SubItems(5) = 0
ic.SubItems(6) = 0
ic.SubItems(7) = 0
ic.SubItems(8) = "0"
ic.Selected = True
Check1.Value = 0
Check2.Value = 1
Check3.Value = 1
Check4.Value = 0
Text5.Text = "0"
List1.ListIndex = 0
Combo1.ListIndex = 0
Combo2.ListIndex = 0
End Sub
Private Sub mnuEditTemplate_Click()
Dim temp As String
On Error GoTo woops
With CommonDialog1
    .DialogTitle = "Replace Menu with Template"
    .CancelError = True
    .Filter = "Menu Template |*.bmu"
    .ShowOpen
    If Len(.FileName) = 0 Then Exit Sub
    temp = .FileName
End With
rtftext.LoadFile temp
curtext = rtftext.Text
ParseMenu
woops:
End Sub
Private Sub mnuFileAbout_Click()
Dim temp As String
temp = "This little App will allow you to edit or create" + vbCrLf + _
"menus in VB6 Forms. New forms can be created with" + vbCrLf + _
"menus in place. It can be used to extract menu" + vbCrLf + _
"structures from one form and place it in another." + vbCrLf + _
"You can save menu structures as templates for later use." + vbCrLf + vbCrLf + _
"It removes the limitations of the number of nested" + vbCrLf + _
"submenus allowable in the VB6 menu editor on which" + vbCrLf + _
"it is based. It has only been tested in VB6." + vbCrLf + vbCrLf + _
"Use it as you would the Menu Editor in VB6 with the" + vbCrLf + _
"exception of the Open/Save operations. As with all" + vbCrLf + _
"my submissions, bugs and errors are provided completely" + vbCrLf + _
"free of charge."
MsgBox temp, vbInformation, "Bobo Enterprises"
End Sub
Private Sub mnuFileExit_Click()
Unload Me
End Sub
Private Sub mnuFileNew_Click()
List1.Clear
LV.ListItems.Clear
Text4.Text = "Form1"
Text6.Text = ""
Label5 = "Form Name :"
ExistingPath = ""
List1.AddItem ""
Set ic = LV.ListItems.Add(, , "")
ic.SubItems(1) = 0
ic.SubItems(2) = 1
ic.SubItems(3) = 1
ic.SubItems(4) = ""
ic.SubItems(5) = 0
ic.SubItems(6) = 0
ic.SubItems(7) = 0
ic.SubItems(8) = "0"
ic.Selected = True
Check1.Value = 0
Check2.Value = 1
Check3.Value = 1
Check4.Value = 0
Text5.Text = "0"
List1.ListIndex = 0
Combo1.ListIndex = 0
Combo2.ListIndex = 0
existing = False
End Sub
Private Sub mnuFileOpen_Click()
'On Error GoTo woops
Dim curtext1 As String
Dim temp As String
Dim temp1 As String
Dim tempInt1 As Integer
Dim tempInt2 As Integer
Dim tempInt3 As Integer
With CommonDialog1
    .DialogTitle = "Edit Existing Form"
    .CancelError = True
    .Filter = "VB 6 Forms |*.frm"
    .ShowOpen
    If Len(.FileName) = 0 Then Exit Sub
    temp = .FileName
End With
Label5 = "Form Name :"
ExistingPath = temp
Text6.Text = temp
existing = True
'read the form to get the befor menu structure and after menu
'structure text and finally the menu structure itself
'We separate it like this to make it easy to put back
'together when we get to saving
rtftext.LoadFile ExistingPath
curtext = rtftext.Text
textfound = InStr(1, curtext, "Attribute VB_Name =")
AfterTxt = Mid(curtext, textfound, Len(curtext) - textfound + 1)
textfound = InStr(1, AfterTxt, vbCrLf)
curtext1 = Left(AfterTxt, textfound)
tempInt1 = InStr(curtext1, Chr(34))
tempInt2 = InStr(tempInt1 + 1, curtext1, Chr(34))
tempInt3 = tempInt2 - tempInt1
temp1 = Mid(curtext1, tempInt1, tempInt3)
temp = Right(temp1, Len(temp1) - 1)
Text4.Text = temp
textfound = InStr(1, curtext, "Begin VB.Menu")
If textfound = -1 Then
    curtext1 = Left(curtext, Len(curtext) - Len(AfterTxt))
    BeforeTxt = Mid$(curtext1, 1, InStrRev(curtext1, "E") - 1)
    Exit Sub
End If
BeforeTxt = Left(curtext, textfound - 1)
curtext = Mid(curtext, Len(BeforeTxt), Len(curtext) - Len(AfterTxt) - Len(BeforeTxt))
ParseMenu
woops:
End Sub
Public Function GetMyMenu() As String
'This function is really messy - but what it does is
'writes to a form or a template the menu structure
'shown in the list, in a format acceptable to VB6
Dim tempstr() As String, emp As String, empcnt() As Integer, diffemp As Integer
Dim chcheck As String, chenable As String, chvis As String, txtIndex As String, cboShcut As String
Dim txtHelpCID As String, chWlist As String, cboNegPos As String
Dim EndCount As Integer, alreadyWlist As Boolean
EndCount = 1
ReDim tempstr(0 To List1.ListCount - 1)
ReDim empcnt(0 To List1.ListCount - 1) 'nested depth
For x = 0 To List1.ListCount - 1
    emp = Mid$(List1.List(x), 1, InStrRev(List1.List(x), "·"))
    empcnt(x) = Len(emp)
    chcheck = ""
    chenable = ""
    chvis = ""
    txtIndex = ""
    cboShcut = ""
    chWlist = ""
    cboNegPos = ""
    txtHelpCID = ""
    'get the data from the hidden ListView
    If LV.ListItems(x + 1).SubItems(1) = 1 Then chcheck = vbCrLf + String(empcnt(x) + 7, " ") + "Checked        =   -1"
    If LV.ListItems(x + 1).SubItems(2) = 0 Then chenable = vbCrLf + String(empcnt(x) + 7, " ") + "Enabled        =   0"
    If LV.ListItems(x + 1).SubItems(3) = 0 Then chvis = vbCrLf + String(empcnt(x) + 7, " ") + "Visible        =   0"
    If LV.ListItems(x + 1).SubItems(4) <> "" Then txtIndex = vbCrLf + String(empcnt(x) + 7, " ") + "Index           =   " + LV.ListItems(x + 1).SubItems(4)
    If LV.ListItems(x + 1).SubItems(5) <> 0 Then cboShcut = vbCrLf + String(empcnt(x) + 7, " ") + "Shortcut        =   " + GetShortCut(Val(LV.ListItems(x + 1).SubItems(5)))
    If LV.ListItems(x + 1).SubItems(6) = 1 Then chWlist = vbCrLf + String(empcnt(x) + 7, " ") + "WindowList      =   -1"
    If LV.ListItems(x + 1).SubItems(7) <> 0 Then cboNegPos = vbCrLf + String(empcnt(x) + 7, " ") + "NegotiatePosition=   " + LV.ListItems(x + 1).SubItems(7)
    If LV.ListItems(x + 1).SubItems(8) = "" Then LV.ListItems(x + 1).SubItems(8) = "0"
    If LV.ListItems(x + 1).SubItems(8) <> "0" Then txtHelpCID = vbCrLf + String(empcnt(x) + 7, " ") + "HelpContextID   =   " + LV.ListItems(x + 1).SubItems(8)
    
    'Make sure the menu structure is valid
    If x = 0 Then
        If empcnt(x) > 0 Then GoTo mnuError1 'read mnuError1 for explanation
    Else
        If empcnt(x) > empcnt(x - 1) + 4 Then GoTo mnuError1
    End If
    If empcnt(x) = 0 Then 'things disallowed in parent menus
        If LV.ListItems(x + 1).SubItems(5) <> 0 Then GoTo mnuError2
        If LV.ListItems(x + 1).SubItems(1) = 1 Then GoTo mnuError3
        If alreadyWlist = True Then
            GoTo mnuError7
        Else
            If LV.ListItems(x + 1).SubItems(6) = 1 Then
                alreadyWlist = True
            End If
        End If
    Else                 'things disallowed in submenus
        If LV.ListItems(x + 1).SubItems(6) = 1 Then GoTo mnuError8
    End If
    'needs a menu name
    If LV.ListItems(x + 1).Text = "" Then GoTo mnuError4
    'make sure everythings OK with index numbers
    If txtIndex = "" Then
        For z = 1 To LV.ListItems.Count
            For p = 1 To LV.ListItems.Count
                If p <> z Then
                    If Len(LV.ListItems(z).SubItems(4)) = 0 Then
                        If Len(LV.ListItems(p).SubItems(4)) = 0 Then
                            If LV.ListItems(z).Text = LV.ListItems(p).Text Then GoTo mnuError5
                        End If
                    End If
                End If
            Next p
        Next z
    Else
        If Val(LV.ListItems(x + 1).SubItems(4)) > 0 Then
            If empcnt(x) <> empcnt(x - 1) Then
                GoTo mnuError6
            Else
                If Val(LV.ListItems(x).SubItems(4)) <> Val(LV.ListItems(x + 1).SubItems(4)) - 1 Then GoTo mnuError6
            End If
        End If
    End If
    'if we get this far the structure must be valid so fill
    'our string array with data
    If x = 0 Then
        tempstr(x) = String(3, " ") + "Begin VB.Menu " + LV.ListItems(x + 1).Text + vbCrLf + String(empcnt(x) + 7, " ") + "Caption        =   " + Chr(34) + Mid$(List1.List(x), InStrRev(List1.List(x), "·") + 1) + Chr(34) + chcheck + chenable + chvis + txtIndex + cboShcut + chWlist + cboNegPos + txtHelpCID
    Else
        If empcnt(x) = empcnt(x - 1) + 4 Then
            tempstr(x) = String(empcnt(x) + 3, " ") + "Begin VB.Menu " + LV.ListItems(x + 1).Text + vbCrLf + String(empcnt(x) + 7, " ") + "Caption        =   " + Chr(34) + Mid$(List1.List(x), InStrRev(List1.List(x), "·") + 1) + Chr(34) + chcheck + chenable + chvis + txtIndex + cboShcut + chWlist + cboNegPos + txtHelpCID
        ElseIf empcnt(x) = empcnt(x - 1) Then
            tempstr(x) = String(empcnt(x - 1) + 3, " ") + "End" + vbCrLf + String(empcnt(x) + 3, " ") + "Begin VB.Menu " + LV.ListItems(x + 1).Text + vbCrLf + String(empcnt(x) + 7, " ") + "Caption        =   " + Chr(34) + Mid$(List1.List(x), InStrRev(List1.List(x), "·") + 1) + Chr(34) + chcheck + chenable + chvis + txtIndex + cboShcut + chWlist + cboNegPos + txtHelpCID
            EndCount = EndCount + 1
        ElseIf empcnt(x) = 0 Then
            tempstr(x) = String(empcnt(x) + 3, " ") + "Begin VB.Menu " + LV.ListItems(x + 1).Text + vbCrLf + String(empcnt(x) + 7, " ") + "Caption        =   " + Chr(34) + Mid$(List1.List(x), InStrRev(List1.List(x), "·") + 1) + Chr(34) + chcheck + chenable + chvis + txtIndex + cboShcut + chWlist + cboNegPos + txtHelpCID
            For Y = 0 To x - EndCount
            tempstr(x) = String(Y * 4 + 3, " ") + "End" + vbCrLf + tempstr(x)
            EndCount = EndCount + 1
            Next Y
        End If
    End If
Next x
'this makes sure we have the right number of 'End' statements
'and to keep it neat, that they are indented appropriately
For x = 0 To List1.ListCount - 1
    If x <> List1.ListCount - 1 Then
        GetMyMenu = GetMyMenu + tempstr(x) + vbCrLf
    Else
        GetMyMenu = GetMyMenu + tempstr(x)
    End If
Next x
diffemp = (List1.ListCount) - EndCount
For Y = diffemp To 1 Step -1
GetMyMenu = GetMyMenu + vbCrLf + String(Y * 4 + 3, " ") + "End"
Next Y
GetMyMenu = GetMyMenu + vbCrLf + String(3, " ") + "End"
Exit Function
'If the menu structure was invalid we end up here
mnuError1:
MsgBox "Menu Item skipped a level"
List1.ListIndex = x
InvalidMenu = True
Exit Function
mnuError2:
MsgBox "Parent Menu cannot have a Shortcut"
List1.ListIndex = x
InvalidMenu = True
Exit Function
mnuError3:
MsgBox "Parent Menu cannot be Checked"
List1.ListIndex = x
InvalidMenu = True
Exit Function
mnuError4:
MsgBox "Menu must have a name"
List1.ListIndex = x
InvalidMenu = True
Exit Function
mnuError5:
MsgBox "Menu name cannot be duplicated"
List1.ListIndex = z
InvalidMenu = True
Exit Function
mnuError6:
MsgBox "Invalid index"
List1.ListIndex = x
InvalidMenu = True
Exit Function
mnuError7:
MsgBox "Only one Window List allowed"
List1.ListIndex = x
InvalidMenu = True
Exit Function
mnuError8:
MsgBox "Only Parent Menu can be a Window List"
List1.ListIndex = x
InvalidMenu = True
Exit Function
End Function
Private Sub mnuFileOpenTemplate_Click()
Dim temp As String
On Error GoTo woops
With CommonDialog1
    .DialogTitle = "Open Template"
    .CancelError = True
    .Filter = "Menu Template |*.bmu"
    .ShowOpen
    If Len(.FileName) = 0 Then Exit Sub
    temp = .FileName
    Text4.Text = .FileTitle
    Text6.Text = .FileName
    Label5 = "Template"
End With
rtftext.LoadFile temp
curtext = rtftext.Text
ParseMenu
woops:
End Sub
Private Sub mnuFileSaveAsTemplate_Click()
Dim temp As String, myMenu As String
On Error GoTo woops
With CommonDialog1
    .DialogTitle = "Save Menu as Template"
    .CancelError = True
    .Filter = "Menu Template |*.bmu"
    .ShowSave
    If Len(.FileName) = 0 Then Exit Sub
    temp = .FileName
End With
Screen.MousePointer = 11
myMenu = GetMyMenu
Screen.MousePointer = 0
FileSave myMenu, temp
woops:
End Sub
Private Sub mnuFileSaveFormAs_Click()
On Error GoTo woops
Dim temp As String, myMenu As String, sfile As String
Screen.MousePointer = 11
myMenu = GetMyMenu
Screen.MousePointer = 0
If InvalidMenu Then
    InvalidMenu = False
    Screen.MousePointer = 0
    Exit Sub
End If
With CommonDialog1
    .FileName = Text4.Text + ".frm"
    .DialogTitle = "Save Form"
    .CancelError = True
    .Filter = "VB 6 Forms |*.frm"
    .ShowSave
    If Len(.FileName) = 0 Then Exit Sub
    sfile = .FileName
End With
temp = BeforeTxt + vbCrLf + myMenu + vbCrLf + "End" + vbCrLf + AfterTxt
FileSave temp, sfile
existing = True
Text6.Text = sfile
ExistingPath = sfile
woops:
Screen.MousePointer = 0
End Sub
Private Sub mnuFileSaveMenu_Click()
On Error GoTo woops
Dim temp As String, myMenu As String, sfile As String
Screen.MousePointer = 11
myMenu = GetMyMenu
Screen.MousePointer = 0
If InvalidMenu Then
    InvalidMenu = False
    Exit Sub
End If
With CommonDialog1
    .FileName = Text4.Text + ".frm"
    .DialogTitle = "Save Form"
    .CancelError = True
    .Filter = "VB 6 Forms |*.frm"
    .ShowSave
    If Len(.FileName) = 0 Then Exit Sub
    sfile = .FileName
End With
temp = "VERSION 5.00" + vbCrLf + "Begin VB.Form " + Text4.Text + vbCrLf + "   Caption         =   " + Chr(34) + Text4.Text + Chr(34) + vbCrLf _
+ "   ClientHeight    =   3195" + vbCrLf + "   ClientLeft      =   60" + vbCrLf + "   ClientTop       =   345" + vbCrLf _
+ "   ClientWidth     =   4680" + vbCrLf + "   LinkTopic       =   " + Chr(34) + "Form1" + Chr(34) + vbCrLf + "   ScaleHeight     =   3195" + vbCrLf _
+ "   ScaleWidth      =   4680" + vbCrLf + "   StartUpPosition =   3" + vbCrLf + myMenu + vbCrLf + "End" + vbCrLf _
+ "Attribute VB_Name = " + Chr(34) + Text4.Text + Chr(34) + vbCrLf + "Attribute VB_GlobalNameSpace = False" + vbCrLf + "Attribute VB_Creatable = False" + vbCrLf _
+ "Attribute VB_PredeclaredId = True" + vbCrLf + "Attribute VB_Exposed = False" + vbCrLf
FileSave temp, sfile
existing = True
Text6.Text = sfile
ExistingPath = sfile
woops:
Screen.MousePointer = 0
End Sub
Private Sub mnuHelp_Click()
mnuFileAbout_Click
End Sub
Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
Dim emp As String
emp = Mid$(List1.List(List1.ListIndex), 1, InStrRev(List1.List(List1.ListIndex), "·"))
List1.List(List1.ListIndex) = emp + Text1.Text
End Sub
Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
LV.ListItems(List1.ListIndex + 1).Text = Text2.Text
End Sub
Private Sub Text3_KeyUp(KeyCode As Integer, Shift As Integer)
LV.ListItems(List1.ListIndex + 1).SubItems(4) = Text3.Text
End Sub
Private Sub Text5_KeyUp(KeyCode As Integer, Shift As Integer)
LV.ListItems(List1.ListIndex + 1).SubItems(8) = Text5.Text
End Sub
Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim emp As String
emp = Mid$(List1.List(List1.ListIndex), 1, InStrRev(List1.List(List1.ListIndex), "·"))
List1.List(List1.ListIndex) = emp + Text1.Text
End Sub
Private Sub Text2_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
LV.ListItems(List1.ListIndex + 1).Text = Text2.Text
End Sub
Private Sub Text3_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
LV.ListItems(List1.ListIndex + 1).SubItems(4) = Text3.Text
End Sub
Public Sub ParseMenu()
'This sub loads an existing menu from either a form
'or a template into the hidden ListView and the
'list used to show the user
Dim x As Integer
Dim newpos As Integer
Dim Blankcnt As Integer
Dim temp As String
Dim temp1 As String
Dim tempInt1 As Integer
Dim tempInt2 As Integer
Dim tempInt3 As Integer
Dim mnuDot As Integer
Dim mnuCount As Integer
List1.Clear
LV.ListItems.Clear
Set ic = LV.ListItems.Add(, , "")
ic.SubItems(1) = 0
ic.SubItems(2) = 1
ic.SubItems(3) = 1
ic.SubItems(4) = ""
ic.SubItems(5) = 0
ic.SubItems(6) = 0
ic.SubItems(7) = 0
mnuCount = 0
mnuDot = 0
pos = 1
Do Until pos >= Len(curtext) - 1
textfound = InStr(pos, curtext, vbCrLf)
If textfound = 0 Then Exit Do
newpos = pos
pos = textfound + 1
temp = Mid(curtext, newpos, pos - newpos)
If InStr(1, temp, "Begin VB.Menu") Then
    'found a menu so load up the hidden Listview
    Blankcnt = InStr(1, temp, "Begin VB.Menu")
    If Blankcnt > 0 Then Blankcnt = Blankcnt - 1
    temp = TrimVoid(Right(temp, Len(temp) - 14 - Blankcnt))
    mnuCount = mnuCount + 1
    LV.ListItems(mnuCount).Text = temp
    LV.ListItems(mnuCount).SubItems(1) = 0
    LV.ListItems(mnuCount).SubItems(2) = 1
    LV.ListItems(mnuCount).SubItems(3) = 1
    LV.ListItems(mnuCount).SubItems(4) = ""
    LV.ListItems(mnuCount).SubItems(5) = 0
    LV.ListItems(mnuCount).SubItems(6) = 0
    LV.ListItems(mnuCount).SubItems(7) = 0
    Set ic = LV.ListItems.Add(, , "")
    ic.SubItems(1) = 0
    ic.SubItems(2) = 1
    ic.SubItems(3) = 1
    ic.SubItems(4) = ""
    ic.SubItems(5) = 0
    ic.SubItems(6) = 0
    ic.SubItems(7) = 0
    GoTo doboy
End If
'read the file for menu data and add if found
'adjusting checks and comboboxes as we go
If InStr(1, temp, "Caption") Then
    Dim intFirstOne As Integer
    Dim intSecondOne As Integer
    Dim intLength As Integer
    temp = Mid$(temp, InStrRev(temp, "=") + 1)
    tempInt1 = InStr(temp, Chr(34))
    tempInt2 = InStr(tempInt1 + 1, temp, Chr(34))
    tempInt3 = tempInt2 - tempInt1
    temp1 = Mid(temp, tempInt1, tempInt3)
    temp = Right(temp1, Len(temp1) - 1)
    If temp = "" Then temp = "-"
    List1.AddItem String(mnuDot * 4, "·") + temp, mnuCount - 1
    List1.Selected(mnuCount - 1) = True
    mnuDot = mnuDot + 1
    GoTo doboy
End If
If InStr(1, temp, "Checked") Then
    LV.ListItems(mnuCount).SubItems(1) = 1
    GoTo doboy
End If
If InStr(1, temp, "Enabled") Then
    LV.ListItems(mnuCount).SubItems(2) = 0
    GoTo doboy
End If
If InStr(1, temp, "Visible") Then
    LV.ListItems(mnuCount).SubItems(3) = 0
    GoTo doboy
End If
If InStr(1, temp, "Index") Then
    temp = TrimVoid(Mid$(temp, InStrRev(temp, "=") + 1))
    LV.ListItems(mnuCount).SubItems(4) = temp
    GoTo doboy
End If
If InStr(1, temp, "Shortcut") Then
    For x = 1 To 79
        temp1 = GetShortCut(x)
        If InStr(1, temp, temp1) Then
            LV.ListItems(mnuCount).SubItems(5) = x
            List1.Column(1, mnuCount - 1) = Combo1.List(x)
            Exit For
        End If
    Next x
    GoTo doboy
End If
If InStr(1, temp, "WindowList") Then
    LV.ListItems(mnuCount).SubItems(6) = 1
    GoTo doboy
End If
If InStr(1, temp, "NegotiatePosition") Then
    temp = TrimVoid(Mid$(temp, InStrRev(temp, "=") + 1))
    LV.ListItems(mnuCount).SubItems(7) = Val(Left(temp, 1))
    GoTo doboy
End If
If InStr(1, temp, "HelpContextID") Then
    temp = TrimVoid(Mid$(temp, InStrRev(temp, "=") + 1))
    LV.ListItems(mnuCount).SubItems(8) = temp
    GoTo doboy
End If
If InStr(1, temp, "End") Then
    mnuDot = mnuDot - 1 'gives us the indented level of the menuitem
    GoTo doboy
End If
doboy:
Loop
List1.ListIndex = 0
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
KeyAscii = 0 'no manually adjusting the path thanks
End Sub
