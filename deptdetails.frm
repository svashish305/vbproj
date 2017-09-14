VERSION 5.00
Begin VB.Form deptdetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payroll Management "
   ClientHeight    =   3930
   ClientLeft      =   1830
   ClientTop       =   2265
   ClientWidth     =   7440
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   7440
   Begin VB.CommandButton cmdlast 
      Height          =   350
      Left            =   6720
      Picture         =   "deptdetails.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3450
      Width           =   495
   End
   Begin VB.CommandButton cmdprev 
      Height          =   350
      Left            =   5760
      Picture         =   "deptdetails.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3450
      Width           =   495
   End
   Begin VB.CommandButton cmdnext 
      Height          =   350
      Left            =   6240
      Picture         =   "deptdetails.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3450
      Width           =   495
   End
   Begin VB.CommandButton cmdfirst 
      Height          =   350
      Left            =   5280
      Picture         =   "deptdetails.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3450
      Width           =   495
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   17
      Top             =   3240
      Width           =   4695
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
         Left            =   3720
         TabIndex        =   8
         Top             =   225
         Width           =   855
      End
      Begin VB.CommandButton cmdexit 
         BackColor       =   &H80000004&
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
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   225
         Width           =   735
      End
      Begin VB.CommandButton cmddel 
         BackColor       =   &H80000004&
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
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   225
         Width           =   735
      End
      Begin VB.CommandButton cmdedit 
         BackColor       =   &H80000004&
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
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   225
         Width           =   735
      End
      Begin VB.CommandButton cmdsave 
         BackColor       =   &H80000004&
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
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   225
         Width           =   735
      End
      Begin VB.CommandButton cmdadd 
         BackColor       =   &H80000004&
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
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   7215
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2280
         TabIndex        =   2
         Text            =   " "
         Top             =   720
         Width           =   3975
      End
      Begin VB.TextBox txt_dptname 
         Height          =   285
         Left            =   2280
         MaxLength       =   25
         TabIndex        =   3
         Text            =   " "
         Top             =   2160
         Width           =   3975
      End
      Begin VB.TextBox txt_dptcode 
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   " "
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Dept Name"
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
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Dept Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   960
         TabIndex        =   15
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Branch Code "
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
         TabIndex        =   14
         Top             =   720
         Width           =   1215
      End
   End
End
Attribute VB_Name = "deptdetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim enableflag As Integer
'------add records to table -----
Private Sub cmdadd_Click()
Dim code, str1, str2, num As String
Dim i, K As Integer

Call clear_all(Me)            'procedure in function

cmdadd.Enabled = False
cmddel.Enabled = False
cmdedit.Enabled = False
cmdexit.Enabled = False
cmdfirst.Enabled = False
cmdnext.Enabled = False
cmdprev.Enabled = False
cmdlast.Enabled = False
cmdsave.Enabled = True
cmdCancel.Enabled = True

If enableflag = 2 Then
    Call txtcmb_disable(Me)        'module procedure
    enableflag = 1
End If
'generate next / new department code
If dpt.EOF = dpt.BOF And dpt.RecordCount < 1 Then
inputagain:
    code = InputBox("Please enter the department code (like :TRN10)", "Payroll:Department Code Generation")
    K = Len(code)
    For i = K To 1 Step -1
        str1 = Mid(code, i, 1)
        If IsNumeric(str1) <> True Then
            str2 = Left(code, i)
            num = Mid(code, i + 1)
            Exit For
        End If
    Next
    If num = "" Then
        MsgBox "Invalid department code. Please enter like : TRN10", , "Payroll:Department Details"
        GoTo inputagain
    Else
        txt_dptcode = code
    End If
Else
    dpt.MoveLast
    code = dpt.Fields(0)
    K = Len(code)
    For i = K To 1 Step -1
        str1 = Mid(code, i, 1)
        If IsNumeric(str1) <> True Then
            str2 = Left(code, i)
            num = Mid(code, i + 1, K)
            Exit For
        End If
    Next
    txt_dptcode = str2 & (CInt(num) + 1)
End If
dpt.AddNew
End Sub
'------delete records from the table---------
Private Sub cmddel_Click()
Dim i As Integer
i = MsgBox("Delete this record ? ", vbQuestion + vbYesNo, "Payroll:Department Details")
If i = vbYes Then
    If Combo1.Text = "" Or txt_dptcode.Text = "" Then
          MsgBox "No record found to delete. " & Chr(13) & "To delete, first select a record.", vbCritical, "Payroll : Delete error"
    Exit Sub
    End If
    dpt.Delete
    MsgBox "Record Deleted"
    Call clear_all(Me)            'procedure in module
    If dpt.RecordCount < 1 And dpt.EOF = dpt.BOF Then
        MsgBox "Zero records in the table now." & Chr(13) & "If required start entering records now", vbInformation, "Payroll : Data Entry"
        Call norec_action
    Else
         Call chk_displayrec
    End If
End If
End Sub
'--------editing records--------
Private Sub cmdedit_Click()
cmdadd.Enabled = False
cmddel.Enabled = False
cmdedit.Enabled = False
cmdexit.Enabled = False
cmdfirst.Enabled = False
cmdnext.Enabled = False
cmdprev.Enabled = False
cmdlast.Enabled = False

cmdsave.Enabled = True
cmdCancel.Enabled = True
dpt.Edit
Call get_dptdata

If enableflag = 2 Then
    Call txtcmb_disable(Me)        'module procedure
    enableflag = 1
End If
End Sub
 '---------exit form----------
Private Sub cmdexit_Click()
Unload deptdetails
Call menu_disable
End Sub

Private Sub cmdfirst_Click()
On Error GoTo err_movfirst
dpt.MoveFirst
Call get_dptdata
Exit Sub
err_movfirst:
    MsgBox "Zero records in the Department Details Table", vbInformation, "Payroll : Data Entry Error"
End Sub

Private Sub cmdlast_Click()
On Error GoTo err_movlast
dpt.MoveLast
Call get_dptdata
Exit Sub
err_movlast:
    MsgBox "Zero records in the Department Details Table", vbInformation, "Payroll : Data Entry Error"
End Sub

Private Sub cmdnext_Click()
On Error GoTo err_movnext
dpt.MoveNext
If dpt.EOF Then
    MsgBox "Current record is the last record", , "Payroll"
    dpt.MovePrevious
    Call get_dptdata
    Exit Sub
Else
    Call get_dptdata
End If
Exit Sub
err_movnext:
     MsgBox "Zero records in the Department Details Table", vbInformation, "Payroll : Data Entry Error"
End Sub

Private Sub cmdprev_Click()
On Error GoTo err_movprev
dpt.MovePrevious
If dpt.BOF Then
    MsgBox "Current record is the first record", , "Payroll"
    dpt.MoveNext
    Call get_dptdata
    Exit Sub
Else
    Call get_dptdata
End If
Exit Sub
err_movprev:
     MsgBox "Zero records in the Department Details Table", vbInformation, "Payroll : Data Entry Error"
End Sub
'---------save record--------
Private Sub cmdsave_Click()
If txt_dptname.Text = "" Then
     MsgBox "Department name cannot be empty. Enter a name for the department and then click on Save", vbCritical, "Payroll : Data entry error"
     txt_dptname.SetFocus
     Exit Sub
End If
cmdadd.Enabled = True
cmddel.Enabled = True
cmdedit.Enabled = True
cmdexit.Enabled = True
cmdfirst.Enabled = True
cmdnext.Enabled = True
cmdprev.Enabled = True
cmdlast.Enabled = True
cmdsave.Enabled = False
cmdCancel.Enabled = False
Call put_dptdata
dpt.Update
Call chk_displayrec
If enableflag = 1 Then
    Call txtcmb_disable(Me)        'module procedure
    enableflag = 2
End If
End Sub
Private Sub Combo1_LostFocus()
If Combo1.ListIndex = 0 Or Combo1.Text = "" Then
    MsgBox "You have not selected the company code or you have selected the heading." & Chr(13) & "Select the values and not the heading.", vbCritical, "Payroll :Data entry error"
    Combo1.Text = ""
    Combo1.SetFocus
End If
End Sub

'------set form default properties -------------
Private Sub Form_Load()
Dim frgkey_status As Integer
Me.Top = 750
Me.Left = 1500

Set db = OpenDatabase(App.Path & "\payroll_db.mdb", , False)
Set brh = db.OpenRecordset("branch", dbOpenDynaset)
frgkey_status = chk_rscount("branch")
Set dpt = db.OpenRecordset("department", dbOpenDynaset)
If frgkey_status < 1 Then
     MsgBox "Zero records found in Branch Table" & Chr(13) & "First add records to the Branch Table and then start adding records to the department table.", vbCritical, "Payroll :Data entry error"
     
     Call disableall(Me)            'procedure in module
     
     cmdexit.Enabled = True
     Exit Sub
Else
     '====if brh contains 0 records then only add and exit buttons should be enabled========
     If dpt.RecordCount < 1 Then
          Call norec_action
     Else
           Call chk_displayrec
     End If
     txt_dptcode.Enabled = False
     cmdsave.Enabled = False
     cmdCancel.Enabled = False
     
     enableflag = 1                    'set flag value for enabling text boxes
     Call txtcmb_disable(Me)        'module procedure
     enableflag = 2
     
     Call get_brhcode
End If
Call menu_disable
End Sub
 '--------extract branch code from branch table-----
Public Sub get_brhcode()
brh.MoveFirst
Combo1.AddItem "Branch Code       Name"
Do While brh.EOF = False
    Combo1.AddItem brh.Fields(0) & "   " & brh.Fields(2)
    brh.MoveNext
Loop
End Sub
'-------get dept details from dept table ------
Public Sub get_dptdata()
With dpt
    txt_dptcode = .Fields(0)
    Combo1.Text = .Fields(1)
    txt_dptname = .Fields(2)
End With
End Sub
'--------put dept details from dept table -------
Public Sub put_dptdata()
Dim str, str1, str2 As String
Dim i As Integer
 '---------to extract branch code from the combo box
str = Combo1.Text
For i = 1 To Len(str) Step 1
    str1 = Mid(str, i, 1)
    If str1 = " " Then
        str2 = Left(str, i)
        Exit For
    End If
Next
With dpt
    .Fields(0) = UCase(txt_dptcode)
    If str2 = "" Then
        .Fields(1) = UCase(str)
    Else
        .Fields(1) = UCase(str2)
    End If
    .Fields(2) = UCase(txt_dptname)
End With
End Sub
'----cancel updations-----------
Private Sub cmdCancel_Click()
dpt.CancelUpdate
MsgBox "Update cancelled"
Call chk_displayrec
cmdadd.Enabled = True
cmddel.Enabled = True
cmdedit.Enabled = True
cmdexit.Enabled = True
cmdfirst.Enabled = True
cmdnext.Enabled = True
cmdprev.Enabled = True
cmdlast.Enabled = True
cmdsave.Enabled = False
cmdCancel.Enabled = False
If enableflag = 1 Then
     Call txtcmb_disable(Me)        'module procedure
     enableflag = 2
End If
End Sub

'Currently no records in the department table
'Allow user only to add record / exit form
Private Sub norec_action()
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
    If dpt.RecordCount > 0 Or dpt.BOF <> dpt.EOF Then
        dpt.MoveFirst
        Call get_dptdata
    End If
End Sub
