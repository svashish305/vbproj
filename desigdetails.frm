VERSION 5.00
Begin VB.Form desigdetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payroll Management System "
   ClientHeight    =   3810
   ClientLeft      =   2085
   ClientTop       =   1875
   ClientWidth     =   7710
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   7710
   Begin VB.CommandButton cmdlast 
      Height          =   375
      Left            =   6480
      Picture         =   "desigdetails.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3330
      Width           =   495
   End
   Begin VB.CommandButton cmdprev 
      Height          =   375
      Left            =   5520
      Picture         =   "desigdetails.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3330
      Width           =   495
   End
   Begin VB.CommandButton cmdnext 
      Height          =   375
      Left            =   6000
      Picture         =   "desigdetails.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3330
      Width           =   495
   End
   Begin VB.CommandButton cmdfirst 
      Height          =   375
      Left            =   5040
      Picture         =   "desigdetails.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3330
      Width           =   495
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   14
      Top             =   3120
      Width           =   4215
      Begin VB.CommandButton cmdcancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3270
         TabIndex        =   8
         Top             =   225
         Width           =   855
      End
      Begin VB.CommandButton cmdexit 
         BackColor       =   &H80000000&
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   225
         Width           =   630
      End
      Begin VB.CommandButton cmddel 
         BackColor       =   &H80000000&
         Caption         =   "Del"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   225
         Width           =   630
      End
      Begin VB.CommandButton cmdedit 
         BackColor       =   &H80000000&
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1380
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   225
         Width           =   630
      End
      Begin VB.CommandButton cmdsave 
         BackColor       =   &H80000000&
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   750
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   225
         Width           =   630
      End
      Begin VB.CommandButton cmdadd 
         BackColor       =   &H80000000&
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   225
         Width           =   630
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   7455
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2880
         TabIndex        =   2
         Text            =   " "
         Top             =   480
         Width           =   3255
      End
      Begin VB.TextBox txt_dsgname 
         Height          =   285
         Left            =   2880
         TabIndex        =   3
         Text            =   " "
         Top             =   2130
         Width           =   3255
      End
      Begin VB.TextBox txt_dsgcode 
         Height          =   285
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   " "
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Department Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   17
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Designation Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   16
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Designation Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   15
         Top             =   1320
         Width           =   1695
      End
   End
End
Attribute VB_Name = "desigdetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim enableflag As Integer
'-------add records to the database-----------
Private Sub cmdadd_Click()
Dim code, str, alpa, num  As String
Dim i, K As Integer
If enableflag = 2 Then
     Call txtcmb_disable(Me)       'Procedure in module
     enableflag = 1
End If
Call clear_all(Me)            'Procedure in module
cmdadd.Enabled = False
cmdedit.Enabled = False
cmddel.Enabled = False
cmdexit.Enabled = False
cmdfirst.Enabled = False
cmdnext.Enabled = False
cmdprev.Enabled = False
cmdlast.Enabled = False
cmdsave.Enabled = True
cmdCancel.Enabled = True
'generate new / next designation code
If dsg.EOF = dsg.BOF And dsg.RecordCount < 1 Then
inputagain:
     code = InputBox("Please input a starting Designation code:(like A1)", "Designation details:Code Generation")
     For i = Len(code) To 1 Step -1
          str = Mid(code, i, 1)
          If IsNumeric(str) <> True Then
               alpa = Left(code, i)
               num = Mid(code, i + 1, Len(code) - 1)
               Exit For
          End If
     Next i
     If num = "" Then
          MsgBox "Invalid Code. Please enter code like : A1", vbCritical + vbOKOnly, "Payroll"
          GoTo inputagain
     End If
     txt_dsgcode.Text = code
Else
     dsg.MoveLast
     code = dsg.Fields(0)
     For i = Len(code) To 1 Step -1
          str = Mid(code, i, 1)
          If IsNumeric(str) <> True Then
               alpa = Left(code, i)
               num = Mid(code, i + 1, Len(code) - 1)
               Exit For
          End If
     Next i
     txt_dsgcode.Text = alpa & (CInt(num) + 1)
End If
dsg.AddNew
End Sub

Private Sub cmdCancel_Click()

If enableflag = 1 Then
     Call txtcmb_disable(Me)       'module procedure
     enableflag = 2
End If
dsg.CancelUpdate
MsgBox "Update cancelled. "
Call chk_displayrec
cmdadd.Enabled = True
cmdedit.Enabled = True
cmddel.Enabled = True
cmdexit.Enabled = True
cmdfirst.Enabled = True
cmdnext.Enabled = True
cmdprev.Enabled = True
cmdlast.Enabled = True
cmdsave.Enabled = False
cmdCancel.Enabled = False
End Sub
'---------delete the records in the database---------
Private Sub cmddel_Click()
Dim i As Integer
i = MsgBox("Delete this record ? ", vbQuestion + vbYesNo, "Payroll")
If i = vbYes Then
      If Combo1.Text = "" Or txt_dsgcode.Text = "" Then
          MsgBox "No record found to delete. " & Chr(13) & "To delete, first select a record.", vbCritical, "Payroll : Delete error"
          Exit Sub
     End If
     dsg.Delete
     Call clear_all(Me)       'procedure in module
     If dsg.RecordCount < 1 Or dsg.EOF = dsg.BOF Then
          MsgBox "Zero records in the table now. If required, you can add records now"
          Call norec_action
          Exit Sub
     Else
          Call chk_displayrec
     End If
End If
End Sub
'---------edit the records in the database------------
Private Sub cmdedit_Click()
Call getdata
If enableflag = 2 Then
     Call txtcmb_disable(Me)       'module procedure
     enableflag = 1
End If

dsg.Edit
cmdadd.Enabled = False
cmdedit.Enabled = False
cmddel.Enabled = False
cmdexit.Enabled = False
cmdfirst.Enabled = False
cmdnext.Enabled = False
cmdprev.Enabled = False
cmdlast.Enabled = False
cmdsave.Enabled = True
cmdCancel.Enabled = True
End Sub
Private Sub cmdexit_Click()
Call menu_disable
Unload Me
End Sub

Private Sub cmdfirst_Click()
On Error GoTo err_movfirst
dsg.MoveFirst
Call getdata
Exit Sub
err_movfirst:
    MsgBox "Zero records in the Designation Details Table", vbInformation, "Payroll : Data Entry Error"
End Sub

Private Sub cmdlast_Click()
On Error GoTo err_movlast
dsg.MoveLast
Call getdata
Exit Sub
err_movlast:
    MsgBox "Zero records in the Designation Details Table", vbInformation, "Payroll : Data Entry Error"
End Sub

Private Sub cmdnext_Click()
On Error GoTo err_movnext
dsg.MoveNext
If dsg.EOF Then
    MsgBox "Current record is the last record", , "Payroll"
    dsg.MovePrevious
    Call getdata
    Exit Sub
Else
    Call getdata
End If
Exit Sub
err_movnext:
    MsgBox "Zero records in the Designation Details Table", vbInformation, "Payroll : Data Entry Error"
End Sub

Private Sub cmdprev_Click()
On Error GoTo err_movprev
dsg.MovePrevious
If dsg.BOF Then
    MsgBox "Current record is the first record", , "Payroll"
    dsg.MoveNext
    Call getdata
    Exit Sub
Else
    Call getdata
End If
Exit Sub
err_movprev:
     MsgBox "Zero records in the Designation Details Table", vbInformation, "Payroll : Data Entry Error"
End Sub
'---------save the user input to the backend----------
Private Sub cmdsave_Click()
If txt_dsgname.Text = "" Then
     MsgBox "Designation name cannot be empty." & Chr(13) & "Enter a value for designation name and then Click on Save. ", vbCritical, "Payroll : Data entry error"
     txt_dsgname.SetFocus
     Exit Sub
End If
cmdadd.Enabled = True
cmdedit.Enabled = True
cmddel.Enabled = True
cmdexit.Enabled = True
cmdfirst.Enabled = True
cmdnext.Enabled = True
cmdprev.Enabled = True
cmdlast.Enabled = True
cmdsave.Enabled = False
cmdCancel.Enabled = False
Call putdata
dsg.Update
If enableflag = 1 Then
     Call txtcmb_disable(Me)       'module procedure
     enableflag = 2
End If
Call chk_displayrec
End Sub

Private Sub Combo1_LostFocus()
If Combo1.ListIndex = 0 Then
    MsgBox "You have selected the heading." & Chr(13) & "Select the values and not the heading.", vbCritical, "Payroll :Data entry error"
    Combo1.Text = ""
    Combo1.SetFocus
End If
End Sub

'---------default form settings -----------------
Private Sub Form_Load()
Dim frgkey_status As String
Me.Top = 500
Me.Left = 1750
Set db = OpenDatabase(App.Path & "\payroll_db.mdb", , False)
Set dpt = db.OpenRecordset("department", dbOpenDynaset)
frgkey_status = chk_rscount("department")
Set dsg = db.OpenRecordset("designation", dbOpenDynaset)
If frgkey_status < 1 Then
     MsgBox "No records found in Department Table" & Chr(13) & "Cannot open Designation details for data entry", vbCritical, "Payroll :Data entry error"
     Call disableall(Me)           'procedure in module
     cmdexit.Enabled = True
     Exit Sub
Else
     Call getdeptinfo
    '====if brh contains 0 records then only add and exit buttons should be enabled========
    If dsg.RecordCount < 1 Then
          Call norec_action
          Exit Sub
    Else
          Call chk_displayrec
          enableflag = 1
          Call txtcmb_disable(Me)        'module procedure
          enableflag = 2
          cmdsave.Enabled = False
          cmdCancel.Enabled = False
     End If
End If
Call menu_disable
End Sub
'------------append values from backend to the form ------
Public Sub getdata()
With dsg
     txt_dsgcode.Text = .Fields(0)
     Combo1.Text = .Fields(1)
     txt_dsgname.Text = .Fields(2)
End With
End Sub
'-----fetch values from the backend to the textboxes----------
Public Sub putdata()
Dim i As Integer
Dim str, str1, str2 As String
str = Combo1.Text
For i = 1 To Len(str) Step 1
     str1 = Mid(str, i, 1)
     If str1 = " " Then
          str2 = Left(str, i)
          Exit For
     End If
Next i
With dsg
     .Fields(0) = UCase(txt_dsgcode.Text)
     If str2 = "" Then
        .Fields(1) = UCase(str)
     Else
        .Fields(1) = UCase(str2)
     End If
     .Fields(2) = UCase(txt_dsgname.Text)
End With
End Sub
'---get department code from dept table
Public Sub getdeptinfo()
     dpt.MoveFirst
     Combo1.AddItem "Department.Code       Name"
     Do While dpt.EOF = False
          Combo1.AddItem (dpt(0) & "      " & dpt(2))
          dpt.MoveNext
     Loop
End Sub

'-------if Currently no records in the department table
''-------Allow user only to add record / exit form
Public Sub norec_action()
    cmdsave.Enabled = False
    cmdedit.Enabled = False
    cmddel.Enabled = False
    cmdCancel.Enabled = False
    cmdfirst.Enabled = False
    cmdprev.Enabled = False
    cmdnext.Enabled = False
    cmdlast.Enabled = False
End Sub
Public Sub chk_displayrec()
    If dsg.RecordCount > 0 Or dsg.BOF <> dsg.EOF Then
        dsg.MoveFirst
        Call getdata
    End If
End Sub


