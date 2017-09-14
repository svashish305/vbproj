VERSION 5.00
Begin VB.Form grddetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payroll Management System "
   ClientHeight    =   3810
   ClientLeft      =   2340
   ClientTop       =   2775
   ClientWidth     =   6810
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   6810
   Begin VB.CommandButton cmdlast 
      Height          =   375
      Left            =   6195
      Picture         =   "grddetails.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3360
      Width           =   465
   End
   Begin VB.CommandButton cmdprev 
      Height          =   375
      Left            =   5265
      Picture         =   "grddetails.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3360
      Width           =   465
   End
   Begin VB.CommandButton cmdnext 
      Height          =   375
      Left            =   5730
      Picture         =   "grddetails.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3360
      Width           =   465
   End
   Begin VB.CommandButton cmdfirst 
      Height          =   375
      Left            =   4800
      Picture         =   "grddetails.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3360
      Width           =   465
   End
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
      Left            =   2940
      TabIndex        =   8
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton cmdexit 
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
      Left            =   3795
      TabIndex        =   0
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton cmddel 
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
      Left            =   2205
      TabIndex        =   7
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton cmdedit 
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
      Left            =   1590
      TabIndex        =   6
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton cmdsave 
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
      Left            =   855
      TabIndex        =   5
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton cmdadd 
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
      Left            =   240
      TabIndex        =   4
      Top             =   3360
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   6495
      Begin VB.TextBox txt_grdname 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2640
         TabIndex        =   3
         Top             =   2280
         Width           =   3375
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2640
         TabIndex        =   2
         Top             =   600
         Width           =   3375
      End
      Begin VB.TextBox txt_grdcode 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2640
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   1
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Grade Name "
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
         Left            =   600
         TabIndex        =   16
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Grade Code"
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
         Left            =   600
         TabIndex        =   15
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label1 
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
         Left            =   600
         TabIndex        =   14
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Height          =   625
      Left            =   120
      TabIndex        =   17
      Top             =   3120
      Width           =   4425
   End
End
Attribute VB_Name = "grddetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim enableflag As Integer

Private Sub cmdadd_Click()
Dim code, str1, str2, num As String
Dim i, K As Integer

Call clear_all(Me)       'procedure in module
If enableflag = 2 Then
    Call txtcmb_disable(Me)     'procedure in module
    enableflag = 1
End If
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
'generate new / next grade code
If grd.EOF = grd.BOF And grd.RecordCount < 1 Then
inputagain:
    code = InputBox("Please enter the Grade Code (like :G1)", "Payroll : Grade Code Generation")
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
        MsgBox "Invalid Grade code. Please enter like : G1", , "Payroll:Grade Details"
        GoTo inputagain
    Else
        txt_grdcode = code
    End If
Else
    grd.MoveLast
    code = grd.Fields(0)
    K = Len(code)
    For i = K To 1 Step -1
        str1 = Mid(code, i, 1)
        If IsNumeric(str1) <> True Then
            str2 = Left(code, i)
            num = Mid(code, i + 1)
            Exit For
        End If
    Next
    txt_grdcode = str2 & (CInt(num) + 1)
End If
grd.AddNew
End Sub

Private Sub cmdCancel_Click()

grd.CancelUpdate
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
    Call txtcmb_disable(Me)     'procedure in module
    enableflag = 2
End If
End Sub

Private Sub cmddel_Click()
Dim i As Integer
i = MsgBox("Delete this record ? ", vbYesNo, "Payroll:GradeDetails")
If i = vbYes Then
     If Combo1.Text = "" Or txt_grdname.Text = "" Then
        MsgBox "No record found to delete. " & Chr(13) & "To delete, first select a record.", vbCritical, "Payroll : Delete error"
        Exit Sub
     End If
     grd.Delete
     MsgBox "Record Deleted."
     Call clear_all(Me)
     If grd.BOF = grd.EOF Or grd.RecordCount < 1 Then
        MsgBox "Zero records in the table now.           " & Chr(13) & "If required start entering records now", vbInformation, "Payroll : Data Entry"
        Call norec_action
     Else
        Call chk_displayrec
     End If
End If
End Sub

Private Sub cmdedit_Click()
Call getdata
grd.Edit
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
    Call txtcmb_disable(Me)     'procedure in module
    enableflag = 1
End If
End Sub

Private Sub cmdexit_Click()
Call menu_disable
Unload Me
End Sub

Private Sub cmdfirst_Click()
On Error GoTo err_movfirst
grd.MoveFirst
Call getdata
Exit Sub
err_movfirst:
       MsgBox "Zero records in the Grade Details Table", vbInformation, "Payroll : Data Entry Error"
End Sub

Private Sub cmdlast_Click()
On Error GoTo err_movlast
grd.MoveLast
Call getdata
Exit Sub
err_movlast:
       MsgBox "Zero records in the Grade Details Table", vbInformation, "Payroll : Data Entry Error"
End Sub

Private Sub cmdnext_Click()
On Error GoTo err_movnext
grd.MoveNext
If grd.EOF Then
    MsgBox "Current record is the last record", , "Payroll"
    grd.MovePrevious
    Call getdata
    Exit Sub
Else
    Call getdata
End If
Exit Sub
err_movnext:
     MsgBox "Zero records in the Grade Details Table", vbInformation, "Payroll : Data Entry Error"
End Sub

Private Sub cmdprev_Click()
On Error GoTo err_movprev
grd.MovePrevious
If grd.BOF Then
    MsgBox "Current record is the first record", , "Payroll"
    grd.MoveNext
    Call getdata
    Exit Sub
Else
    Call getdata
End If
Exit Sub
err_movprev:
     MsgBox "Zero records in the Grade Details Table", vbInformation, "Payroll : Data Entry Error"
End Sub

Private Sub cmdsave_Click()
If txt_grdname = "" Then
    MsgBox "Grade name cannot be empty." & Chr(13) & "First enter a value for grade name and then Click on Save", vbCritical, "Payroll : Data entry error"
    txt_grdname.SetFocus
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
grd.Update
If enableflag = 1 Then
   Call txtcmb_disable(Me)     'procedure in module
    enableflag = 2
End If
Call chk_displayrec
End Sub

Private Sub Combo1_LostFocus()
If Combo1.ListIndex = 0 Or Combo1.Text = "" Then
    MsgBox "Either you have not selected the designation code or you have selected the heading." & Chr(13) & "Select a valid designation code / name.", vbCritical, "Payroll :Data entry error"
    Combo1.Text = ""
    Combo1.SetFocus
End If
End Sub

Private Sub Form_Load()
Dim frgkey_status As Integer
Me.Top = 1500
Me.Left = 2500

Set db = OpenDatabase(App.Path & "\payroll_db.mdb", , False)
Set dsg = db.OpenRecordset("designation", dbOpenDynaset)
frgkey_status = chk_rscount("designation")
Set grd = db.OpenRecordset("grade", dbOpenDynaset)
If frgkey_status < 1 Then
     MsgBox "No records found in Master Designation Table" & Chr(13) & "Cannot open Grade details for data entry", vbCritical, "Payroll :Data entry error"
     Call disableall(Me)      'procedure in module
     cmdexit.Enabled = True
     Exit Sub
Else
    Call get_dsgdetails
    '====if brh contains 0 records then only add and exit buttons should be enabled========
    If grd.RecordCount < 1 Then
        Call norec_action
    Else
        Call chk_displayrec
        enableflag = 1
        Call txtcmb_disable(Me)     'procedure in module
        enableflag = 2
        txt_grdcode.Enabled = False
        cmdsave.Enabled = False
        cmdCancel.Enabled = False
    End If
End If
Call menu_disable
End Sub

Public Sub get_dsgdetails()
dsg.MoveFirst
Combo1.AddItem "Desg.Code       Name"
Do While dsg.EOF = False
    Combo1.AddItem dsg.Fields(0) & "      " & dsg.Fields(2)
    dsg.MoveNext
Loop
End Sub

Public Sub getdata()
With grd
    txt_grdcode = .Fields(0)
    Combo1.Text = .Fields(1)
    txt_grdname = .Fields(2)
End With
End Sub

Public Sub putdata()
Dim i As Integer
Dim str, str1, str2 As String
str = Combo1.Text
For i = 1 To Len(str)
    str1 = Mid(str, i, 1)
    If str1 = " " Then
        str2 = Left(str, i)
        Exit For
    End If
Next
With grd
    .Fields(0) = UCase(txt_grdcode)
    If str2 = "" Then
        .Fields(1) = UCase(str)
    Else
        .Fields(1) = UCase(str2)
    End If
    .Fields(2) = UCase(txt_grdname)
End With
End Sub
'-------if Currently no records in the department table
'-------Allow user only to add record / exit form
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
    If grd.RecordCount > 0 Or grd.BOF <> grd.EOF Then
        grd.MoveFirst
        Call getdata
    End If
End Sub

