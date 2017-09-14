VERSION 5.00
Begin VB.Form leave_master 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payroll Management System "
   ClientHeight    =   5760
   ClientLeft      =   1830
   ClientTop       =   2130
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   7800
   Begin VB.CommandButton cmdlast 
      Height          =   350
      Left            =   6960
      Picture         =   "leave_master.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5160
      Width           =   495
   End
   Begin VB.CommandButton cmdnext 
      Height          =   350
      Left            =   6480
      Picture         =   "leave_master.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5160
      Width           =   495
   End
   Begin VB.CommandButton cmdprev 
      Height          =   350
      Left            =   6000
      Picture         =   "leave_master.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5160
      Width           =   495
   End
   Begin VB.CommandButton cmdfirst 
      Height          =   350
      Left            =   5520
      Picture         =   "leave_master.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5160
      Width           =   495
   End
   Begin VB.Frame Frame3 
      Height          =   650
      Left            =   240
      TabIndex        =   25
      Top             =   4920
      Width           =   5055
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
         Left            =   3300
         TabIndex        =   12
         Top             =   240
         Width           =   915
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
         Left            =   4215
         TabIndex        =   13
         Top             =   240
         Width           =   735
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
         Left            =   2445
         TabIndex        =   11
         Top             =   240
         Width           =   855
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
         TabIndex        =   10
         Top             =   240
         Width           =   855
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
         TabIndex        =   9
         Top             =   240
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
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4695
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      Begin VB.ComboBox Yrcombo 
         Height          =   315
         ItemData        =   "leave_master.frx":1108
         Left            =   1800
         List            =   "leave_master.frx":110A
         TabIndex        =   2
         Top             =   840
         Width           =   1335
      End
      Begin VB.ComboBox empcombo 
         Height          =   315
         Left            =   1800
         TabIndex        =   1
         Text            =   " "
         Top             =   360
         Width           =   3855
      End
      Begin VB.Frame Frame2 
         Caption         =   "Leave Breakup"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   1080
         TabIndex        =   19
         Top             =   1440
         Width           =   3855
         Begin VB.TextBox txt_oth 
            Height          =   285
            Left            =   2400
            TabIndex        =   7
            Text            =   " "
            Top             =   2400
            Width           =   855
         End
         Begin VB.TextBox txt_ml 
            Height          =   285
            Left            =   2400
            TabIndex        =   6
            Text            =   " "
            Top             =   1920
            Width           =   855
         End
         Begin VB.TextBox txt_pl 
            Height          =   285
            Left            =   2400
            TabIndex        =   5
            Text            =   " "
            Top             =   1440
            Width           =   855
         End
         Begin VB.TextBox txt_sk 
            Height          =   285
            Left            =   2400
            TabIndex        =   4
            Text            =   " "
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox txt_cs 
            Height          =   285
            Left            =   2400
            TabIndex        =   3
            Text            =   " "
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label6 
            Caption         =   "Other Leave"
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
            TabIndex        =   24
            Top             =   2400
            Width           =   1335
         End
         Begin VB.Label Label5 
            Caption         =   "Maternity Leave"
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
            TabIndex        =   23
            Top             =   1920
            Width           =   1575
         End
         Begin VB.Label Label4 
            Caption         =   "Privilege Leave"
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
            TabIndex        =   22
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label Label3 
            Caption         =   "Sick Leave"
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
            TabIndex        =   21
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "Casual Leave"
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
            TabIndex        =   20
            Top             =   480
            Width           =   1455
         End
      End
      Begin VB.Label Label7 
         Caption         =   "Year"
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
         Left            =   1200
         TabIndex        =   26
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Employee Code"
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
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   1575
      End
   End
End
Attribute VB_Name = "leave_master"
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

cmdadd.Enabled = False
cmdedit.Enabled = False
cmddel.Enabled = False
cmdfirst.Enabled = False
cmdprev.Enabled = False
cmdnext.Enabled = False
cmdlast.Enabled = False
cmdsave.Enabled = True
cmdCancel.Enabled = True

empcombo.Enabled = True
levdetails.AddNew
End Sub

Private Sub cmdCancel_Click()
levdetails.CancelUpdate
If enableflag = 1 Then
    Call enable
    enableflag = 2
End If

cmdadd.Enabled = True
cmdedit.Enabled = True
cmddel.Enabled = True
cmdfirst.Enabled = True
cmdprev.Enabled = True
cmdnext.Enabled = True
cmdlast.Enabled = True
cmdsave.Enabled = False
cmdCancel.Enabled = False

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
    levdetails.Delete
    Call cleardata
    MsgBox "deleted"
End If
End Sub

Private Sub cmdedit_Click()
cmdadd.Enabled = False
cmdedit.Enabled = False
cmddel.Enabled = False
cmdfirst.Enabled = False
cmdprev.Enabled = False
cmdnext.Enabled = False
cmdlast.Enabled = False
cmdsave.Enabled = True
cmdCancel.Enabled = True
flag = 1
If enableflag = 2 Then
    Call enable
    enableflag = 1
End If
'levdetails.Edit
empcombo.Enabled = True
End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdfirst_Click()
    levdetails.MoveFirst
    Call getdata
    MsgBox "First record ", , "Payroll"
End Sub

Private Sub cmdlast_Click()
levdetails.MoveLast
    Call getdata
    MsgBox "Last record ", , "Payroll"
End Sub

Private Sub cmdnext_Click()
levdetails.MoveNext
If levdetails.EOF Then
    MsgBox "Current record is the last record", , "Payroll"
    levdetails.MovePrevious
    Call getdata
    Exit Sub
Else
    Call getdata
End If
End Sub

Private Sub cmdprev_Click()
levdetails.MovePrevious
If levdetails.BOF Then
    MsgBox "Current record is the first record", , "Payroll"
    levdetails.MoveNext
    Call getdata
    Exit Sub
Else
    Call getdata
End If
End Sub

Private Sub cmdsave_Click()
Call putdata
levdetails.Update
If enableflag = 1 Then
    Call enable
    enableflag = 2
End If
cmdadd.Enabled = True
cmdedit.Enabled = True
cmddel.Enabled = True
cmdfirst.Enabled = True
cmdprev.Enabled = True
cmdnext.Enabled = True
cmdlast.Enabled = True
cmdsave.Enabled = False
cmdCancel.Enabled = False

empcombo.Enabled = False
levdetails.MoveFirst
Call getdata
End Sub

Private Sub empcombo_LostFocus()
Dim str As String
Dim lval As String, str1 As String
Dim i As Integer
If flag = 1 Then
str = empcombo.Text
For i = 1 To Len(str) Step 1
    str1 = Mid(str, i, 1)
    If str1 = " " Then
        lval = Left(str, i - 1)
        Exit For
    End If
Next
levdetails.MoveFirst
Do While levdetails.EOF = False
    If levdetails.Fields(0) = lval Then
    getdata
    levdetails.Edit
    flag = 0
    Exit Do
    End If
    levdetails.MoveNext
Loop
End If
End Sub

Private Sub Form_Load()
Me.Top = 1500
Me.Left = 2000
Set db = OpenDatabase(App.Path & "\payroll_db.mdb", , False)
Set levdetails = db.OpenRecordset("emp_leave_details", dbOpenDynaset)

If levdetails.EOF Then
     cmdfirst.Enabled = False
     cmdnext.Enabled = False
     cmdprev.Enabled = False
     cmdlast.Enabled = False
End If

Call getempdata
empcombo.Enabled = False

enableflag = 1
If enableflag = 1 Then
    Call enable
    enableflag = 2
End If

cmdsave.Enabled = False
cmdCancel.Enabled = False
End Sub

Public Sub enable()
     Yrcombo.Enabled = Not Yrcombo.Enabled
     txt_cs.Enabled = Not txt_cs.Enabled
     txt_sk.Enabled = Not txt_sk.Enabled
     txt_pl.Enabled = Not txt_pl.Enabled
     txt_ml.Enabled = Not txt_ml.Enabled
     txt_oth.Enabled = Not txt_oth.Enabled
End Sub

Public Sub getdata()
With levdetails
    empcombo.Text = .Fields(0)
    Yrcombo.Text = .Fields(1)
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
With levdetails
    If lval = "" Then
        .Fields(0) = str
    Else
        .Fields(0) = lval
    End If
    .Fields(1) = Yrcombo.Text
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
empcombo.Text = ""
txt_cs.Text = ""
txt_sk.Text = ""
txt_pl.Text = ""
txt_oth.Text = ""
txt_ml.Text = ""
Yrcombo.Text = ""
End Sub
Private Sub Yrcombo_GotFocus()
Dim yrval, yrcur As Integer
yrcur = Year(Now)
For yrval = 0 To 2 Step 1
     Yrcombo.AddItem (yrcur + yrval)
Next yrval
Yrcombo.Text = yrcur
End Sub
