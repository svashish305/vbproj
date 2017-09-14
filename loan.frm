VERSION 5.00
Begin VB.Form loan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payroll Management System "
   ClientHeight    =   4695
   ClientLeft      =   1575
   ClientTop       =   2130
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   8655
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
      Left            =   3780
      TabIndex        =   13
      Top             =   4200
      Width           =   870
   End
   Begin VB.ComboBox empcombo 
      Height          =   315
      Left            =   2760
      TabIndex        =   2
      Text            =   " "
      Top             =   240
      Width           =   4815
   End
   Begin VB.CommandButton CmdLast 
      Height          =   350
      Left            =   7125
      Picture         =   "loan.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4200
      Width           =   495
   End
   Begin VB.CommandButton CmdNext 
      Height          =   350
      Left            =   6630
      Picture         =   "loan.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4200
      Width           =   495
   End
   Begin VB.CommandButton CmdPrev 
      Height          =   350
      Left            =   6135
      Picture         =   "loan.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4200
      Width           =   495
   End
   Begin VB.CommandButton CmdFirst 
      Height          =   350
      Left            =   5640
      Picture         =   "loan.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4200
      Width           =   495
   End
   Begin VB.CommandButton CmdExit 
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
      Left            =   4650
      TabIndex        =   0
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton CmdDel 
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
      Left            =   2310
      TabIndex        =   11
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton CmdEdit 
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
      Left            =   3045
      TabIndex        =   12
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton CmdSave 
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
      Left            =   1575
      TabIndex        =   10
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton CmdAdd 
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
      TabIndex        =   1
      Top             =   4200
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Loan Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   360
      TabIndex        =   19
      Top             =   720
      Width           =   7935
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6240
         TabIndex        =   30
         Text            =   " "
         Top             =   2010
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2535
         Width           =   1095
      End
      Begin VB.TextBox bal_amt 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6240
         TabIndex        =   7
         Text            =   " "
         Top             =   2535
         Width           =   1095
      End
      Begin VB.TextBox txt_rate 
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   " "
         Top             =   960
         Width           =   735
      End
      Begin VB.ComboBox lncombo 
         Height          =   315
         Left            =   3120
         TabIndex        =   3
         Top             =   360
         Width           =   3375
      End
      Begin VB.TextBox txt_bal 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6240
         TabIndex        =   9
         Text            =   " "
         Top             =   1485
         Width           =   1095
      End
      Begin VB.TextBox txt_instal 
         Height          =   285
         Left            =   6240
         TabIndex        =   8
         Text            =   " "
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txt_davail 
         Height          =   285
         Left            =   2160
         TabIndex        =   4
         Text            =   " "
         Top             =   1485
         Width           =   1095
      End
      Begin VB.TextBox txt_amt 
         Height          =   285
         Left            =   2160
         TabIndex        =   5
         Text            =   " "
         Top             =   2010
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   " Amount / Instalment"
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
         Left            =   4215
         TabIndex        =   29
         Top             =   2010
         Width           =   1770
      End
      Begin VB.Label Label9 
         Caption         =   "Total  Amount to be repay"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   480
         TabIndex        =   28
         Top             =   2535
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Amount repaid"
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
         Left            =   4680
         TabIndex        =   27
         Top             =   2535
         Width           =   1305
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Rate of Interest"
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
         Left            =   600
         TabIndex        =   26
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Balance  Instalments "
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
         Left            =   4095
         TabIndex        =   25
         Top             =   1485
         Width           =   1890
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Total  Instalments "
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
         Left            =   4380
         TabIndex        =   24
         Top             =   960
         Width           =   1605
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Date of availing"
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
         Left            =   540
         TabIndex        =   23
         Top             =   1485
         Width           =   1395
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Amount availed"
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
         Left            =   540
         TabIndex        =   22
         Top             =   2010
         Width           =   1395
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Loan Type "
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
         Left            =   1920
         TabIndex        =   20
         Top             =   360
         Width           =   1020
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      Height          =   240
      Left            =   1080
      TabIndex        =   18
      Top             =   240
      Width           =   1470
   End
End
Attribute VB_Name = "loan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim enableflag As Integer
Dim flag As Integer


Private Sub cmdadd_Click()
Call cleardata
Call cmdenable
empcombo.Enabled = True
empcombo.SetFocus
If enableflag = 2 Then
    Call enable
    enableflag = 1
End If
ln.AddNew
End Sub

Private Sub cmdCancel_Click()
ln.CancelUpdate
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
    If Val(txt_bal.Text) = 0 Then
        ln.Delete
        Call cleardata
        MsgBox "deleted"
    Else
        MsgBox "Delete not allowed. Loan balance to be cleared first,Kindly re-check", vbCritical, "Payroll"
    End If
End If
End Sub

Private Sub cmdedit_Click()
If enableflag = 2 Then
    Call enable
    enableflag = 1
End If
empcombo.Enabled = True
Call cmdenable
ln.Edit
End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdfirst_Click()
ln.MoveFirst
Call getdata
MsgBox "First record ", , "Payroll"
End Sub

Private Sub cmdlast_Click()
ln.MoveLast
Call getdata
MsgBox "Last record ", , "Payroll"
End Sub

Private Sub cmdnext_Click()
ln.MoveNext
If ln.EOF Then
    MsgBox "Current record is the last record", , "Payroll"
    ln.MovePrevious
    Call getdata
    Exit Sub
Else
    Call getdata
End If
End Sub

Private Sub cmdprev_Click()
ln.MovePrevious
If ln.BOF Then
    MsgBox "Current record is the first record", , "Payroll"
    ln.MoveNext
    Call getdata
    Exit Sub
Else
    Call getdata
End If
End Sub

Private Sub cmdsave_Click()
On Error GoTo errsave
Call putdata
ln.Update

If enableflag = 1 Then
    Call enable
    enableflag = 2
End If
Call cmdenable
empcombo.Enabled = False
Exit Sub
errsave:
    MsgBox "Employee already availed this loan. Please re-check on the loan type. ", vbExclamation, "Payroll"
    lncombo.SetFocus
End Sub


Private Sub Form_Load()
Me.Top = 1500
Me.Left = 2000
Set db = OpenDatabase(App.Path & "\payroll_db.mdb", , False)
Set ln = db.OpenRecordset("Loan_availed", dbOpenDynaset)
Set lm = db.OpenRecordset("loan_master", dbOpenDynaset)
enableflag = 1
If enableflag = 1 Then
    Call enable
    enableflag = 2
End If
cmdsave.Enabled = False
cmdCancel.Enabled = False
Call getempdata
Call get_lm
empcombo.Enabled = False
End Sub

Public Sub enable()
lncombo.Enabled = Not lncombo.Enabled
txt_amt.Enabled = Not txt_amt.Enabled
txt_rate.Enabled = Not txt_rate.Enabled
bal_amt.Enabled = Not bal_amt.Enabled
txt_davail.Enabled = Not txt_davail.Enabled
txt_instal.Enabled = Not txt_instal.Enabled
txt_bal.Enabled = Not txt_bal.Enabled
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
With ln
    empcombo.Text = .Fields(0)
    lncombo.Text = .Fields(1)
    txt_davail.Text = .Fields(2)
    txt_amt.Text = .Fields(3)
    bal_amt.Text = .Fields(4)
    txt_instal.Text = .Fields(5)
    txt_bal.Text = .Fields(6)
    Text2.Text = .Fields(7)
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
With ln
    If lval = " " Then
        .Fields(0) = str
    Else
        .Fields(0) = lval
    End If
    .Fields(1) = lncombo.Text
    .Fields(2) = txt_davail.Text
    .Fields(3) = txt_amt.Text
    .Fields(4) = 0
    .Fields(5) = txt_instal.Text
    .Fields(6) = txt_instal.Text
    .Fields(7) = Text2.Text
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
Dim txt As Control
For Each txt In Controls
    If TypeOf txt Is TextBox Then
        txt.Text = ""
    End If
Next
empcombo.Text = ""
lncombo.Text = ""
End Sub

Public Sub get_lm()
Set loan_query = db.OpenRecordset("select * from loan_master")
loan_query.MoveFirst
Do While loan_query.EOF = False
    lncombo.AddItem loan_query.Fields(0)
    loan_query.MoveNext
Loop
End Sub

Private Sub lncombo_LostFocus()
Set loan_rate = db.OpenRecordset("select rate_interest from loan_master where loan_master.ln_name= '" & lncombo.Text & "'", dbOpenDynaset)
txt_rate.Text = loan_rate.Fields(0)
End Sub

Private Sub txt_amt_LostFocus()
Text1.Text = Val(txt_amt) + ((Val(txt_rate) / 100) * Val(txt_amt))
End Sub

Private Sub txt_instal_LostFocus()
Text2.Text = Val(Text1.Text) / Val(txt_instal.Text)
End Sub
