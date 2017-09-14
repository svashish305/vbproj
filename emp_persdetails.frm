VERSION 5.00
Begin VB.Form emp_pers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payroll Management System "
   ClientHeight    =   5985
   ClientLeft      =   1275
   ClientTop       =   525
   ClientWidth     =   8910
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   8910
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
      Height          =   350
      Left            =   840
      TabIndex        =   21
      Top             =   5520
      Width           =   710
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
      Height          =   350
      Left            =   3660
      TabIndex        =   0
      Top             =   5520
      Width           =   710
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
      Height          =   350
      Left            =   2955
      TabIndex        =   24
      Top             =   5520
      Width           =   710
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
      Height          =   350
      Left            =   2250
      TabIndex        =   23
      Top             =   5520
      Width           =   710
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
      Height          =   350
      Left            =   1545
      TabIndex        =   22
      Top             =   5520
      Width           =   710
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
      Height          =   350
      Left            =   4365
      TabIndex        =   25
      Top             =   5520
      Width           =   855
   End
   Begin VB.CommandButton cmdlast 
      Height          =   350
      Left            =   7440
      Picture         =   "emp_persdetails.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   5520
      Width           =   480
   End
   Begin VB.CommandButton cmdprev 
      Height          =   350
      Left            =   6480
      Picture         =   "emp_persdetails.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   5520
      Width           =   480
   End
   Begin VB.CommandButton cmdnext 
      Height          =   345
      Left            =   6960
      Picture         =   "emp_persdetails.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   5520
      Width           =   480
   End
   Begin VB.CommandButton cmdfirst 
      Height          =   350
      Left            =   6000
      Picture         =   "emp_persdetails.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5520
      Width           =   480
   End
   Begin VB.Frame Frame2 
      Height          =   3855
      Left            =   120
      TabIndex        =   31
      Top             =   1440
      Width           =   8655
      Begin VB.CommandButton cmd_salary 
         BackColor       =   &H00C0C0FF&
         Caption         =   "&SALARY"
         Height          =   315
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   3360
         Width           =   1695
      End
      Begin VB.TextBox txt_tel 
         Height          =   285
         Left            =   1800
         TabIndex        =   13
         Text            =   " "
         Top             =   3000
         Width           =   2775
      End
      Begin VB.TextBox txt_pin 
         Height          =   285
         Left            =   1800
         TabIndex        =   12
         Text            =   " "
         Top             =   2640
         Width           =   2775
      End
      Begin VB.TextBox txt_addr3 
         Height          =   285
         Left            =   1800
         TabIndex        =   11
         Text            =   " "
         Top             =   2400
         Width           =   2775
      End
      Begin VB.TextBox txt_addr2 
         Height          =   285
         Left            =   1800
         TabIndex        =   10
         Text            =   " "
         Top             =   2160
         Width           =   2775
      End
      Begin VB.TextBox txt_djoin 
         Height          =   285
         Left            =   7200
         TabIndex        =   14
         Text            =   " "
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txt_dconfirm 
         Height          =   285
         Left            =   7200
         TabIndex        =   15
         Text            =   " "
         Top             =   720
         Width           =   1215
      End
      Begin VB.ComboBox bl_grpcombo 
         Height          =   315
         ItemData        =   "emp_persdetails.frx":1108
         Left            =   6720
         List            =   "emp_persdetails.frx":1124
         TabIndex        =   19
         Text            =   "bl_grpcombo"
         Top             =   2755
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Married ( Yes / No)"
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
         Left            =   5280
         TabIndex        =   18
         Top             =   2160
         Width           =   2175
      End
      Begin VB.TextBox txt_addr1 
         Height          =   285
         Left            =   1800
         TabIndex        =   9
         Text            =   " "
         Top             =   1920
         Width           =   2775
      End
      Begin VB.TextBox txt_dob 
         Height          =   285
         Left            =   1800
         TabIndex        =   7
         Top             =   1050
         Width           =   975
      End
      Begin VB.TextBox txt_qualif 
         Height          =   285
         Left            =   1800
         TabIndex        =   8
         Text            =   " "
         Top             =   1455
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Female"
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
         Left            =   6840
         TabIndex        =   17
         Top             =   1440
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Male"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5520
         TabIndex        =   16
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txt_father 
         Height          =   285
         Left            =   1800
         TabIndex        =   6
         Text            =   " "
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox txt_empname 
         Height          =   285
         Left            =   1800
         TabIndex        =   5
         Top             =   240
         Width           =   2775
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sex"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5280
         TabIndex        =   35
         Top             =   1200
         Width           =   2775
      End
      Begin VB.Label Label15 
         Caption         =   "Telephone"
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
         Left            =   120
         TabIndex        =   45
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "PinCode"
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
         Left            =   120
         TabIndex        =   44
         Top             =   2685
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Blood Group"
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
         Left            =   5280
         TabIndex        =   43
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "Date of Joining"
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
         Left            =   5280
         TabIndex        =   42
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "Date of Confirmation"
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
         Left            =   5280
         TabIndex        =   41
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label10 
         Caption         =   "Date Of Birth"
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
         Left            =   120
         TabIndex        =   40
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Qualification"
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
         Left            =   120
         TabIndex        =   39
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Father's Name"
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
         Left            =   120
         TabIndex        =   38
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Address"
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
         Left            =   120
         TabIndex        =   37
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Employee Name"
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
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   30
      Top             =   120
      Width           =   8655
      Begin VB.ComboBox gradecombo 
         Height          =   315
         ItemData        =   "emp_persdetails.frx":115A
         Left            =   6480
         List            =   "emp_persdetails.frx":115C
         TabIndex        =   4
         Top             =   240
         Width           =   1815
      End
      Begin VB.ComboBox dsgcombo 
         Height          =   315
         Left            =   3600
         TabIndex        =   3
         Text            =   " "
         Top             =   240
         Width           =   1935
      End
      Begin VB.ComboBox deptcombo 
         Height          =   315
         Left            =   720
         TabIndex        =   2
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txt_empcode 
         Height          =   285
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label16 
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
         Height          =   495
         Left            =   5880
         TabIndex        =   46
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Desg. Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3000
         TabIndex        =   34
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Dept. Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   33
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Emp.Code"
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
         Left            =   2520
         TabIndex        =   32
         Top             =   840
         Width           =   975
      End
   End
End
Attribute VB_Name = "emp_pers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim enableflag As Integer
Dim cmb_enable As Integer
Dim cmdval As Integer   'status of request generated by employee personal
Dim edit_sal_status As Integer

Private Sub cmd_salary_Click()
salary_master.Show
emp_pers.Hide
If cmdval = 1 Then
'Adding a new employee record, salary - master to be opened for add
    Set sal_mast = db.OpenRecordset("salary_master", dbOpenDynaset)
    salary_master.cmdokay.Enabled = True
    salary_master.cmdExit.Enabled = False
    salary_master.cleardata
    salary_master.tgross.Text = 0
    sal_mast.AddNew
    salary_master.tbasic.SetFocus
ElseIf cmdval = 0 Then
'viewing employee records, salary master to be opened for viewing
    Set sal_mast = db.OpenRecordset("select * from salary_master where emp_code = '" & txt_empcode.Text & "'", dbOpenDynaset)
    salary_master.cmdokay.Enabled = False
    salary_master.cmdExit.Enabled = True
    salary_master.getdata
    salary_master.enabletxt
    salary_master.cmdExit.SetFocus
ElseIf cmdval = 2 Then
'Editing an employee record, salary master to be opened in edit
    edit_sal_status = 1
    emp_pers.cmd_salary.Enabled = True
    Set sal_mast = db.OpenRecordset("select * from salary_master where emp_code = '" & txt_empcode.Text & "'", dbOpenDynaset)
    salary_master.cmdokay.Enabled = True
    salary_master.cmdExit.Enabled = False
    salary_master.getdata
    sal_mast.Edit
    salary_master.tbasic.SetFocus
ElseIf cmdval = 3 Then
'deleting an employee record, salary master to be opened in delete mode
    Set sal_mast = db.OpenRecordset("select * from salary_master where emp_code = '" & txt_empcode.Text & "'", dbOpenDynaset)
    salary_master.cmdokay.Enabled = True
    salary_master.cmdExit.Enabled = False
    salary_master.getdata
    sal_mast.Delete
End If
End Sub
'------add record (s) to the database--------
Private Sub cmdadd_Click()
Dim code, str, num, alp As String
Dim i, K  As Integer
'-------------
cmdval = 1

Call clear_all(Me)            'procedure in module

cmdadd.Enabled = False
cmddel.Enabled = False
cmdedit.Enabled = False
cmdExit.Enabled = False
cmdfirst.Enabled = False
cmdnext.Enabled = False
cmdprev.Enabled = False
cmdlast.Enabled = False
cmdsave.Enabled = True
cmdCancel.Enabled = True
If enableflag = 2 Then
    Call txt_disable(Me)         'procedure in module
    Call combo_enable
    enableflag = 1
End If
' for automatic generation of branch code:user's choice
If emp.EOF = emp.BOF And emp.RecordCount < 1 Then
chkcode:
    code = InputBox("Please enter the employee code (like PSG100):", "Payroll : Employee Code Generation")
    K = Len(code)
    For i = K To 1 Step -1
        str = Mid(code, i, 1)
        If IsNumeric(str) <> True Then
            alp = Mid(code, 1, i)
            num = Mid(code, i + 1)
            Exit For
        End If
    Next i
    If num = "" Then
        MsgBox "Invalid value for employee code, Please enter like 'PSG100'. ", , "Payroll:Branch Details"
        GoTo chkcode
    Else
        txt_empcode = code
    End If
Else
    emp.MoveLast
    code = emp(0)
    K = Len(code)
    For i = K To 1 Step -1
        str = Mid(code, i, 1)
        If IsNumeric(str) <> True Then
            alp = Mid(code, 1, i)
            num = Mid(code, i + 1)
            Exit For
        End If
    Next i
        txt_empcode = alp & (CInt(num) + 1)
End If
emp_pers.cmd_salary.Enabled = True
emp.AddNew
End Sub
'---------cancel update / edit ----------
Private Sub cmdCancel_Click()
emp.CancelUpdate
MsgBox "Update record cancelled"
Call chk_displayrec
cmdadd.Enabled = True
cmdedit.Enabled = True
cmddel.Enabled = True
cmdExit.Enabled = True
cmdfirst.Enabled = True
cmdnext.Enabled = True
cmdprev.Enabled = True
cmdlast.Enabled = True
cmdsave.Enabled = False
cmdCancel.Enabled = False
If enableflag = 1 Then
    Call txt_disable(Me)         'procedure in module
    Call combo_enable
    enableflag = 2
End If

deptcombo.Enabled = False
dsgcombo.Enabled = False
gradecombo.Enabled = False

emp_pers.cmd_salary.Enabled = False
End Sub
'-------delete record (s) to the table ---------
Private Sub cmddel_Click()
Dim i As Integer
i = MsgBox("Delete this record ? ", vbQuestion + vbYesNo, "Payroll")
If i = vbYes Then
    If txt_empcode.Text = "" Or deptcombo.Text = "" Then
        MsgBox "No record found to delete. " & Chr(13) & "To delete, first select a record.", vbCritical, "Payroll : Delete error"
        Exit Sub
    Else
        emp_pers.cmd_salary.Enabled = False
        cmdval = 3
        emp.Delete
        Call clear_all(Me)         'procedure in module
        If emp.RecordCount < 1 Or emp.EOF Or emp.BOF Then
            Call chk_displayrec
        End If
    End If
End If
End Sub
'-----------edit current record -----------
Private Sub cmdedit_Click()
Call getdata
emp.Edit
'-------------
cmdval = 2
emp_pers.cmd_salary.Enabled = True
cmdadd.Enabled = False
cmdedit.Enabled = False
cmddel.Enabled = False
cmdExit.Enabled = False
cmdfirst.Enabled = False
cmdnext.Enabled = False
cmdprev.Enabled = False
cmdlast.Enabled = False
cmdsave.Enabled = True
cmdCancel.Enabled = True
If enableflag = 2 Then
    Call txt_disable(Me)     'procedure in module
    Call combo_enable
    txt_empcode.Enabled = False
   ' deptcombo.SetFocus
    enableflag = 1
End If
End Sub
'------exit working on the form ---------
Private Sub cmdexit_Click()
Call menu_disable
Unload Me
End Sub
'-----------move to first record --------
Private Sub cmdfirst_Click()
On Error GoTo err_movfirst
cmdval = 0
If emp.BOF <> True Then
    emp.MoveFirst
    Call getdata
    emp_pers.cmd_salary.Enabled = True
End If
Exit Sub
err_movfirst:
    MsgBox "Zero records in the Employee Details Table", vbInformation, "Payroll : Data Entry Error"
End Sub
'---------move to last record --------
Private Sub cmdlast_Click()
On Error GoTo err_movlast
cmdval = 0
If emp.EOF <> True Then
    emp.MoveLast
    Call getdata
    emp_pers.cmd_salary.Enabled = True
End If
Exit Sub
err_movlast:
    MsgBox "Zero records in the Employee Details Table", vbInformation, "Payroll : Data Entry Error"
End Sub
'----------move to the next record -------
Private Sub cmdnext_Click()
On Error GoTo err_movnext
cmdval = 0
emp.MoveNext
If emp.EOF Then
    MsgBox "Current record is the last record", , "Payroll"
    emp.MovePrevious
    Call getdata
    emp_pers.cmd_salary.Enabled = True
    Exit Sub
End If
Call getdata
Exit Sub
err_movnext:
    MsgBox "Zero records in the Employee Details Table", vbInformation, "Payroll : Data Entry Error"
End Sub
'---------move to the previous record ----------
Private Sub cmdprev_Click()
On Error GoTo err_movprevious
cmdval = 0
emp.MovePrevious
If emp.BOF Then
    MsgBox "Current record is the first record", , "Payroll"
    emp.MoveNext
    Call getdata
    emp_pers.cmd_salary.Enabled = True
    Exit Sub
End If
Call getdata
Exit Sub
err_movprevious:
    MsgBox "Zero records in the Employee Details Table", vbInformation, "Payroll : Data Entry Error"
End Sub
'---------save the added / edited record --------
Private Sub cmdsave_Click()
cmdadd.Enabled = True
cmddel.Enabled = True
cmdedit.Enabled = True
cmdExit.Enabled = True
cmdfirst.Enabled = True
cmdnext.Enabled = True
cmdprev.Enabled = True
cmdlast.Enabled = True
cmdsave.Enabled = False
cmdCancel.Enabled = False
If enableflag = 1 Then
    Call txt_disable(Me)     'module procedure
    Call combo_enable
    enableflag = 2
End If
Call putdata
emp.Update
If cmdval = 2 Then
    If edit_sal_status = 1 Then
        sal_mast.Update
        'GoTo cont
    End If
End If

MsgBox " Salary details saved.", vbInformation, "Payroll : Save record"
Unload salary_master
cmd_salary.Enabled = False
Call chk_displayrec
End Sub

Private Sub deptcombo_GotFocus()
dpt.MoveFirst
deptcombo.Clear
deptcombo.AddItem "Dept.code    Name"
Do While dpt.EOF = False
     deptcombo.AddItem dpt.Fields(0)
     dpt.MoveNext
Loop
End Sub
Private Sub deptcombo_LostFocus()
If deptcombo.ListIndex = 0 Then
     MsgBox "You have selected the heading." & Chr(13) & "Select the Department code, name and not the heading", vbCritical, "Payroll : Data entry error"
     deptcombo.Text = ""
     deptcombo.SetFocus
     Exit Sub
Else
     'deptcombo.Enabled = False
     dsgcombo.Clear
     dsgcombo.Enabled = True
     dsgcombo.AddItem "Desg.code   Name"
     Set dept_dsg = db.OpenRecordset("select dsg_code from designation where dpt_code='" & deptcombo.Text & "'", dbOpenDynaset)
End If
End Sub
Private Sub dsgcombo_GotFocus()
If dept_dsg.EOF Then
     MsgBox "No designations found under this department"
     deptcombo.SetFocus
     Exit Sub
Else
    deptcombo.Enabled = False
     dept_dsg.MoveFirst
     Do While dept_dsg.EOF = False
          dsgcombo.AddItem dept_dsg.Fields(0)
          dept_dsg.MoveNext
     Loop
End If
End Sub
Private Sub dsgcombo_LostFocus()
If dsgcombo.ListIndex = 0 Then
     MsgBox "You have selected the heading." & Chr(13) & "Select the Designation code, name and not the heading", vbCritical, "Payroll : Data entry error"
     dsgcombo.Text = ""
     dsgcombo.SetFocus
     Exit Sub
Else
     'dsgcombo.Enabled = False
     gradecombo.Clear
     gradecombo.AddItem "Grade.code     Name"
     gradecombo.Enabled = True
     Set dsg_grd = db.OpenRecordset("select grd_code from grade where dsg_code = '" & dsgcombo.Text & "'", dbOpenDynaset)
End If
End Sub
'----------default settings in form load---------
Private Sub Form_Load()
Dim frgkey_status As Integer

Me.Top = 250
Me.Left = 1500

Set db = OpenDatabase(App.Path & "\payroll_db.mdb", , False)
Set dpt = db.OpenRecordset("department", dbOpenDynaset)
Set dsg = db.OpenRecordset("designation", dbOpenDynaset)
Set grd = db.OpenRecordset("grade", dbOpenDynaset)

frgkey_status = chk_rscount("grade")
Set emp = db.OpenRecordset("emp_personal", dbOpenDynaset)
If frgkey_status < 1 Then
    MsgBox "Zero records in the Grade Table" & Chr(13) & "First add records to the Company Table and then start adding records to the Grade Table. ", vbCritical, "Payroll :Data entry error"
    Call disableall(Me)            'procedure in module
    cmdExit.Enabled = True
    Exit Sub
End If

'====if brh contains 0 records then only add and exit buttons should be enabled========
If emp.RecordCount < 1 Then
    Call norec_action
Else
    Call chk_displayrec
    enableflag = 1
    Call txt_disable(Me)     'module procedure
    Call combo_enable
    cmd_salary.Enabled = False
    enableflag = 2
    cmdsave.Enabled = False
    cmdCancel.Enabled = False
    cmdval = 0
End If
Call menu_disable
End Sub
'----------get data from the backend to the form-------
Public Sub getdata()
With emp
    deptcombo.Text = .Fields(16)
    dsgcombo.Text = .Fields(17)
    gradecombo.Text = .Fields(1)
    txt_empcode.Text = .Fields(0)
    txt_empname.Text = .Fields(2)
    txt_father.Text = .Fields(3)
    txt_dob.Text = .Fields(4)
    txt_addr1.Text = .Fields(5)
    txt_addr2.Text = .Fields(6)
    txt_addr3.Text = .Fields(7)
    txt_pin.Text = .Fields(8)
    txt_tel.Text = .Fields(9)
    If .Fields(10) = True Then      '-------get values for marital status
        Check1.Value = Checked
    Else
        Check1.Value = Unchecked
    End If
    If .Fields(11) = True Then      '----get values for sex type
        Option1.Value = True
    Else
        Option2.Value = True
    End If
    bl_grpcombo.Text = .Fields(12)       '-----get value for blood group
    txt_qualif.Text = .Fields(13)
    txt_djoin.Text = .Fields(14)
    txt_dconfirm.Text = .Fields(15)
End With
End Sub
'---------save data to the backend --------
Public Sub putdata()
With emp
    .Fields(0) = UCase(txt_empcode.Text)
    .Fields(16) = UCase(deptcombo.Text)
    .Fields(17) = UCase(dsgcombo.Text)
    .Fields(1) = UCase(gradecombo.Text)
    .Fields(2) = UCase(txt_empname.Text)
    .Fields(3) = UCase(txt_father.Text)
    .Fields(4) = Format(txt_dob.Text, "dd/mm/yy")
    .Fields(5) = UCase(txt_addr1.Text)
    .Fields(6) = UCase(txt_addr2.Text)
    .Fields(7) = UCase(txt_addr3.Text)
    .Fields(8) = Val(txt_pin.Text)
    .Fields(9) = UCase(txt_tel.Text) & " "
    If Check1.Value = True Then     '---append values for marital status
        .Fields(10) = "True"
    Else
        .Fields(10) = "False"
    End If
inputsex:
    If Option1.Value = True Then
        .Fields(11) = "True"
    ElseIf Option2.Value = True Then
        .Fields(11) = "False"
    Else
        MsgBox "Insufficient data. Please click for Employee Sex type", vbQuestion, "Payroll:Employee Details"
        GoTo inputsex
    End If
    .Fields(12) = UCase(bl_grpcombo.Text)
    .Fields(13) = UCase(txt_qualif.Text)
    .Fields(14) = Format(txt_djoin.Text, "dd/mm/yy")
    .Fields(15) = Format(txt_dconfirm.Text, "dd/mm/yy")
End With
End Sub
Private Sub gradecombo_GotFocus()
If dsg_grd.EOF Then
     MsgBox "No Grades found under this designation"
     dsgcombo.SetFocus
     Exit Sub
Else
    dsgcombo.Enabled = False
    dsg_grd.MoveFirst
    Do While dsg_grd.EOF = False
        gradecombo.AddItem dsg_grd.Fields(0)
        dsg_grd.MoveNext
    Loop
End If
End Sub
Private Sub gradecombo_LostFocus()
If gradecombo.ListIndex = 0 Then
     MsgBox "You have selected the heading." & Chr(13) & "Select the grade code, name and not the heading", vbCritical, "Payroll : Data entry error"
     gradecombo.Text = ""
     gradecombo.SetFocus
     Exit Sub
Else
     gradecombo.Enabled = False
End If
End Sub
'Status of command buttons in form if no records found in employee table
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
'Check whether valid date entered or not
Private Sub txt_dconfirm_LostFocus()
    If IsDate(txt_dconfirm.Text) = False Then
        MsgBox "You have entered a wrong value for date." & Chr(13) & "Valid date format :'dd/mm/yy'", vbCritical, "Payroll : Data entry error"
        txt_dconfirm.Text = ""
       ' txt_dconfirm.SetFocus
    End If
End Sub
'Check whether valid date entered or not
Private Sub txt_djoin_LOSTFOCUS()
    If IsDate(txt_djoin.Text) = False Then
        MsgBox "You have entered a wrong value for date." & Chr(13) & "Valid date format :'dd/mm/yy'", vbCritical, "Payroll : Data entry error"
        txt_djoin.Text = ""
    End If
End Sub
'Check whether valid date entered or not
Private Sub txt_dob_LostFocus()
    If IsDate(txt_dob.Text) = False Then
        MsgBox "You have entered a wrong value for date." & Chr(13) & "Valid date format :'dd/mm/yy'", vbCritical, "Payroll : Data entry error"
        txt_dob.Text = ""
    End If
End Sub

Public Sub chk_displayrec()
    If emp.RecordCount > 0 Or emp.BOF <> emp.EOF Then
        emp.MoveFirst
        cmdval = 0
        Call getdata
    End If
End Sub

Public Sub combo_enable()
    deptcombo.Enabled = Not deptcombo.Enabled
    dsgcombo.Enabled = Not dsgcombo.Enabled
    gradecombo.Enabled = Not gradecombo.Enabled
End Sub

