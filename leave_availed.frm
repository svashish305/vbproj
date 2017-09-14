VERSION 5.00
Begin VB.Form leave_availed 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payroll Management System "
   ClientHeight    =   5370
   ClientLeft      =   1830
   ClientTop       =   1995
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   7470
   Begin VB.ComboBox monthcombo 
      Height          =   315
      ItemData        =   "leave_availed.frx":0000
      Left            =   3360
      List            =   "leave_availed.frx":0028
      TabIndex        =   1
      Text            =   " "
      Top             =   960
      Width           =   1695
   End
   Begin VB.ComboBox empcombo 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "leave_availed.frx":008E
      Left            =   1680
      List            =   "leave_availed.frx":0090
      TabIndex        =   0
      Text            =   " "
      Top             =   240
      Width           =   4935
   End
   Begin VB.CommandButton cmdlast 
      Height          =   375
      Left            =   6720
      Picture         =   "leave_availed.frx":0092
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4680
      Width           =   495
   End
   Begin VB.CommandButton cmdnext 
      Height          =   375
      Left            =   6240
      Picture         =   "leave_availed.frx":04D4
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4680
      Width           =   495
   End
   Begin VB.CommandButton cmdprev 
      Height          =   375
      Left            =   5760
      Picture         =   "leave_availed.frx":0916
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4680
      Width           =   495
   End
   Begin VB.CommandButton cmdfirst 
      Height          =   375
      Left            =   5280
      Picture         =   "leave_availed.frx":0D58
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4680
      Width           =   495
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   240
      TabIndex        =   29
      Top             =   4560
      Width           =   4890
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
         Left            =   3120
         TabIndex        =   11
         Top             =   210
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
         Left            =   3990
         TabIndex        =   12
         Top             =   210
         Width           =   750
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
         Left            =   2370
         TabIndex        =   10
         Top             =   210
         Width           =   750
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
         Left            =   1620
         TabIndex        =   9
         Top             =   210
         Width           =   750
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
         Left            =   870
         TabIndex        =   8
         Top             =   210
         Width           =   750
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
         Left            =   120
         TabIndex        =   7
         Top             =   210
         Width           =   750
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "LEAVE"
      Height          =   3615
      Left            =   240
      TabIndex        =   22
      Top             =   720
      Width           =   6975
      Begin VB.Frame Frame4 
         Caption         =   "BALANCE"
         Height          =   2775
         Left            =   3600
         TabIndex        =   31
         Top             =   600
         Width           =   2895
         Begin VB.TextBox bal_oth 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1680
            TabIndex        =   21
            Text            =   " "
            Top             =   2280
            Width           =   975
         End
         Begin VB.TextBox bal_ml 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1680
            TabIndex        =   20
            Text            =   " "
            Top             =   1800
            Width           =   975
         End
         Begin VB.TextBox bal_pl 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1680
            TabIndex        =   19
            Text            =   " "
            Top             =   1320
            Width           =   975
         End
         Begin VB.TextBox bal_sk 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1680
            TabIndex        =   18
            Text            =   " "
            Top             =   840
            Width           =   975
         End
         Begin VB.TextBox bal_cs 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1680
            TabIndex        =   17
            Text            =   " "
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label13 
            Caption         =   "Other Leave"
            Height          =   255
            Left            =   240
            TabIndex        =   37
            Top             =   2280
            Width           =   1095
         End
         Begin VB.Label Label12 
            Caption         =   "Maternity Leave"
            Height          =   255
            Left            =   240
            TabIndex        =   36
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label Label11 
            Caption         =   "Privilege Leave"
            Height          =   375
            Left            =   240
            TabIndex        =   35
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label10 
            Caption         =   "Sick Leave"
            Height          =   255
            Left            =   240
            TabIndex        =   34
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label9 
            Caption         =   "Casual Leave"
            Height          =   255
            Left            =   240
            TabIndex        =   33
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "AVAILED"
         Height          =   2775
         Left            =   360
         TabIndex        =   23
         Top             =   600
         Width           =   2775
         Begin VB.TextBox txt_oth 
            Height          =   285
            Left            =   1560
            TabIndex        =   6
            Text            =   " "
            Top             =   2280
            Width           =   975
         End
         Begin VB.TextBox txt_ml 
            Height          =   285
            Left            =   1560
            TabIndex        =   5
            Text            =   " "
            Top             =   1800
            Width           =   975
         End
         Begin VB.TextBox txt_pl 
            Height          =   285
            Left            =   1560
            TabIndex        =   4
            Text            =   " "
            Top             =   1320
            Width           =   975
         End
         Begin VB.TextBox txt_sk 
            Height          =   285
            Left            =   1560
            TabIndex        =   3
            Text            =   " "
            Top             =   840
            Width           =   975
         End
         Begin VB.TextBox txt_cs 
            Height          =   285
            Left            =   1560
            TabIndex        =   2
            Text            =   " "
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label7 
            Caption         =   "Other Leave"
            Height          =   255
            Left            =   480
            TabIndex        =   28
            Top             =   2280
            Width           =   975
         End
         Begin VB.Label Label6 
            Caption         =   "Maternity Leave"
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label Label5 
            Caption         =   "Privilege Leave"
            Height          =   255
            Left            =   240
            TabIndex        =   26
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "Sick Leave"
            Height          =   255
            Left            =   480
            TabIndex        =   25
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "Casual Leave"
            Height          =   255
            Left            =   360
            TabIndex        =   24
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Month"
         Height          =   195
         Left            =   2400
         TabIndex        =   32
         Top             =   240
         Width           =   450
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Employee Code"
      Height          =   255
      Left            =   480
      TabIndex        =   30
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "leave_availed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim enableflag As Integer
Dim flag As Integer

Private Sub cmdadd_Click()
Call cleardata
If enableflag = 2 Then
    Call enable
    enableflag = 1
End If
Call cmdenable
empcombo.Enabled = True
empcombo.SetFocus
levavailed.AddNew
End Sub

Private Sub cmdCancel_Click()
levavailed.CancelUpdate
If enableflag = 1 Then
    Call enable
    enableflag = 2
End If
Call cmdenable
empcombo.Enabled = False
End Sub

Private Sub cmddel_Click()
Dim i As Integer
i = MsgBox("Delete this record ? ", vbQuestion + vbYesNo, "Payroll")
If i = vbYes Then
    If empcombo.Text = " " Then
        MsgBox "No record found to delete.", vbCritical, "Payroll"
        Exit Sub
    End If
    levavailed.Delete
    Call cleardata
    MsgBox "deleted"
End If
End Sub

Private Sub cmdedit_Click()
Call cmdenable
flag = 1
If enableflag = 2 Then
    Call enable
    enableflag = 1
End If
levavailed.Edit
'empcombo.Enabled = True
End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdfirst_Click()
levavailed.MoveFirst
Call getdata

Call getbal_leave
End Sub

Private Sub cmdlast_Click()
levavailed.MoveLast
Call getdata

Call getbal_leave
End Sub

Private Sub cmdnext_Click()
levavailed.MoveNext
If levavailed.EOF Then
    MsgBox "Current record is the last record", , "Payroll"
    levavailed.MovePrevious
    Call getdata
    Call getbal_leave
    Exit Sub
Else
    Call getdata
    Call getbal_leave
    
End If
End Sub

Private Sub cmdprev_Click()
levavailed.MovePrevious
If levavailed.BOF Then
    MsgBox "Current record is the first record", , "Payroll"
    levavailed.MoveNext
    Call getdata
    Call getbal_leave
    
    Exit Sub
Else
    Call getdata
    Call getbal_leave
    
End If
End Sub

Private Sub cmdsave_Click()
On Error GoTo errsave
Call putdata
levavailed.Update
leavedetup
If enableflag = 1 Then
    Call enable
    enableflag = 2
End If
Call cmdenable
empcombo.Enabled = False
Exit Sub
errsave:
    MsgBox "Duplicate entry for the employee for the same month. Kindly check", vbExclamation, "Payroll"
    empcombo.SetFocus
End Sub

'Private Sub empcombo_LostFocus()
'Dim str As String
'Dim lval As String, str1 As String
'Dim i As Integer
'If flag = 1 Then
'str = empcombo.Text
'For i = 1 To Len(str) Step 1
'    str1 = Mid(str, i, 1)
'    If str1 = " " Then
'        lval = Left(str, i - 1)
'        Exit For
'    End If
'Next
'levavailed.MoveFirst
'Do While levavailed.EOF = False
'    If levavailed.Fields(0) = lval Then
'    getdata
'    levavailed.Edit
'    flag = 0
'    Exit Do
'    End If
'    levavailed.MoveNext
'Loop
'End If
'End Sub

Private Sub Form_Load()
Me.Top = 1500
Me.Left = 2000
Set db = OpenDatabase(App.Path & "\payroll_db.mdb", , False)
Set levavailed = db.OpenRecordset("emp_leave_availed", dbOpenDynaset)


'If levavailed.EOF Then
'    cmdfirst.Enabled = False
'    cmdnext.Enabled = False
'    cmdprev.Enabled = False
'    cmdlast.Enabled = False
'End If
Call getempdata
enableflag = 1
If enableflag = 1 Then
    Call enable
    enableflag = 2
End If
cmdsave.Enabled = False
cmdCancel.Enabled = False
End Sub

Public Sub enable()
monthcombo.Enabled = Not monthcombo.Enabled
txt_cs.Enabled = Not txt_cs.Enabled
txt_sk.Enabled = Not txt_sk.Enabled
txt_pl.Enabled = Not txt_pl.Enabled
txt_ml.Enabled = Not txt_ml.Enabled
txt_oth.Enabled = Not txt_oth.Enabled
End Sub

Public Sub cmdenable()
Dim cont As Control
For Each cont In Controls
If TypeOf cont Is CommandButton Then
cont.Enabled = Not cont.Enabled
End If
Next
End Sub

Public Sub getdata()
With levavailed
    empcombo.Text = .Fields(0)
    monthcombo.Text = .Fields(1)
    txt_cs.Text = .Fields(2)
    txt_sk.Text = .Fields(3)
    txt_pl.Text = .Fields(4)
    txt_oth.Text = .Fields(5)
    txt_ml.Text = .Fields(6)
End With
End Sub

Public Sub putdata()
Dim str, str1, lval As String
Dim i As Integer
'-------------extract cmpcode from combobox text value
str = empcombo.Text
For i = 1 To Len(str) Step 1
    str1 = Mid(str, i, 1)
    If str1 = " " Then
        lval = Left(str, i - 1)
        Exit For
    End If
Next
With levavailed
    If lval = "" Then
        .Fields(0) = str
    Else
        .Fields(0) = lval
    End If
    .Fields(1) = monthcombo.Text
    .Fields(2) = txt_cs.Text
    .Fields(3) = txt_sk.Text
    .Fields(4) = txt_pl.Text
    .Fields(5) = txt_oth.Text
    .Fields(6) = txt_ml.Text
End With
End Sub

Public Sub getempdata()
Set emp = db.OpenRecordset("emp_personal", dbOpenDynaset)
emp.MoveFirst
Do While emp.EOF = False
    empcombo.AddItem emp.Fields(0) & "     " & emp.Fields(1)
    emp.MoveNext
Loop
End Sub
 
Public Sub cleardata()
Dim clr As Control
For Each clr In Controls
    If TypeOf clr Is TextBox Then
        clr.Text = ""
    End If
Next
empcombo.Text = ""
monthcombo.Text = ""
End Sub

Private Sub empcombo_LostFocus()
Dim str, str1, empstring As String
Dim i As Integer
'-------------extract cmpcode from combobox text value
str = empcombo.Text
For i = 1 To Len(str) Step 1
    str1 = Mid(str, i, 1)
    If str1 = " " Then
        empstring = Left(str, i - 1)
        Exit For
    End If
Next
Set leave_query = db.OpenRecordset("Select * from emp_leave_details where emp_code='" & empstring & "'", dbOpenDynaset)
With leave_query
    bal_cs.Text = .Fields(1)
    bal_sk.Text = .Fields(2)
    bal_pl.Text = .Fields(3)
    bal_oth.Text = .Fields(4)
    bal_ml.Text = .Fields(5)
End With
End Sub

Private Sub txt_cs_LostFocus()
bal_cs.Text = Val(bal_cs.Text) - Val(txt_cs.Text)
End Sub

Private Sub txt_sk_LostFocus()
bal_sk.Text = Val(bal_sk.Text) - Val(txt_sk.Text)
End Sub

Private Sub txt_pl_LostFocus()
bal_pl.Text = Val(bal_pl.Text) - Val(txt_pl.Text)
End Sub

Private Sub txt_oth_LostFocus()
bal_oth.Text = Val(bal_oth.Text) - Val(txt_oth.Text)
End Sub

Private Sub txt_ml_LostFocus()
bal_ml.Text = Val(bal_ml.Text) - Val(txt_ml.Text)
End Sub

Public Sub getbal_leave()
Set leave_query = db.OpenRecordset("Select * from emp_leave_details where emp_code='" & empcombo.Text & "'", dbOpenDynaset)
'leave_query.Refresh
With leave_query
    bal_cs.Text = .Fields(1)
    bal_sk.Text = .Fields(2)
    bal_pl.Text = .Fields(3)
    bal_oth.Text = .Fields(4)
    bal_ml.Text = .Fields(5)
End With
End Sub

Private Sub leavedetup()
leave_query.Edit
With leave_query
      .Fields(1) = bal_cs.Text
     .Fields(2) = bal_sk.Text
    .Fields(3) = bal_pl.Text
    .Fields(4) = bal_oth.Text
     .Fields(5) = bal_ml.Text
End With
leave_query.Update
End Sub

