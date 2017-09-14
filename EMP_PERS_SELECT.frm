VERSION 5.00
Begin VB.Form emprpt_selecttype 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payroll Management System "
   ClientHeight    =   2730
   ClientLeft      =   2475
   ClientTop       =   3285
   ClientWidth     =   7035
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   7035
   Begin VB.CommandButton cmdexit 
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3120
      TabIndex        =   8
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdokay 
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmddept 
      Caption         =   "&Deptwise Employees Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   1320
      Width           =   3735
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   " &Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5760
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdindividual 
      Caption         =   "&Selected Employees Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   1680
      Width           =   3735
   End
   Begin VB.CommandButton cmdall 
      Caption         =   "&All Employees General Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   960
      Width           =   3735
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2520
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   1560
      TabIndex        =   6
      Top             =   720
      Width           =   3975
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
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   1470
   End
End
Attribute VB_Name = "emprpt_selecttype"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdall_Click()
Gen_Employee_pers.Show
End Sub

Private Sub cmdCancel_Click()
cmdall.Visible = True
cmddept.Visible = True
cmdexit.Visible = True
cmdokay.Visible = False
cmdCancel.Visible = False
Label1.Visible = False
Combo1.Visible = False
End Sub

Private Sub cmddept_Click()
Deptwise_emprpt.Show
End Sub

Private Sub cmdexit_Click()
Unload Me
frmrptChoice.Show
End Sub

Private Sub cmdindividual_Click()
cmdokay.Visible = True
cmdCancel.Visible = True
Label1.Visible = True
Combo1.Visible = True
cmdall.Visible = False
cmddept.Visible = False
cmdexit.Visible = False
Call getdata
End Sub

Private Sub cmdokay_Click()
Dim str, str1, lval As String
Dim nam As String
Dim empind As String
Dim i As Integer
'-------------extract empcode from combobox text value
str = Combo1.Text
For i = 1 To Len(str) Step 1
    str1 = Mid(str, i, 1)
    If str1 = " " Then
        lval = Left(str, i - 1)
        nam = Mid(str, i)
        Exit For
    End If
Next
With emp_rpt
    If lval = "" Then
        empind = str
    Else
        empind = lval
    End If
End With

If emprpt = True Then
    Payroll_environ.rsindividual_emp.Open
    emprpt = False
End If
Payroll_environ.rsindividual_emp.Requery (0)
Payroll_environ.rsindividual_emp.Filter = "emp_code='" & empind & "'"
emprpt_individual.Sections("Section4").Controls.Item("lblName").Caption = nam
If Payroll_environ.rsindividual_emp.RecordCount < 1 Then
    MsgBox "No employee records found for report. ", vbInformation, "Payroll"
Else
    emprpt_individual.Show
End If
End Sub

Public Sub getdata()
Set db = OpenDatabase(App.Path & "\payroll_db.mdb", , False)
Set emp_rpt = db.OpenRecordset("select emp_code, emp_name from emp_personal", dbOpenDynaset)
emp_rpt.MoveFirst
Do While emp_rpt.EOF = False
    Combo1.AddItem emp_rpt.Fields(0) & "    " & emp_rpt.Fields(1)
    emp_rpt.MoveNext
Loop
End Sub

Private Sub Form_Load()
Me.Top = 1500
Me.Left = 2000
End Sub

