VERSION 5.00
Begin VB.Form leav_rptfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payroll Management System "
   ClientHeight    =   2970
   ClientLeft      =   3495
   ClientTop       =   3540
   ClientWidth     =   6570
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   6570
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1320
      TabIndex        =   8
      Text            =   " "
      Top             =   1920
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "&Exit"
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
      Left            =   5400
      TabIndex        =   7
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton cmdcurrent 
      Caption         =   "All Employees current leave status"
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
      Left            =   960
      TabIndex        =   6
      Top             =   840
      Width           =   3975
   End
   Begin VB.CommandButton cmdmonth 
      Caption         =   "All Employees Monthwise Report"
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
      Left            =   960
      TabIndex        =   5
      Top             =   480
      Width           =   3975
   End
   Begin VB.CommandButton cmdind 
      Caption         =   "Individual Employee Leave Details"
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
      Left            =   960
      TabIndex        =   4
      Top             =   120
      Width           =   3975
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "&Cancel"
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
      Left            =   3600
      TabIndex        =   3
      Top             =   2400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdokay 
      Caption         =   "&Ok"
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
      Left            =   2040
      TabIndex        =   2
      Top             =   2400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "leav_rptfrm.frx":0000
      Left            =   1440
      List            =   "leav_rptfrm.frx":0028
      TabIndex        =   1
      Top             =   1920
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      Height          =   1575
      Left            =   360
      Top             =   1320
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.Label Label1 
      Caption         =   "Select the Month for which the report is required : "
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
      Left            =   720
      TabIndex        =   0
      Top             =   1440
      Visible         =   0   'False
      Width           =   4935
   End
End
Attribute VB_Name = "leav_rptfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim indrpt As Integer      'indrpt - individual employee leave report requested by user
Dim mnthrpt As Integer     'mnthrpt - employee specified month required
Dim empind As String       'employee code selected by user for report generation

Private Sub cmdCancel_Click()
Shape1.Visible = False
cmdokay.Visible = False
cmdCancel.Visible = False
Label1.Visible = False
Combo1.Text = ""
Combo2.Text = ""
Combo1.Visible = False
Combo2.Visible = False
cmdind.Visible = True
cmdmonth.Visible = True
cmdcurrent.Visible = True
cmdexit.Visible = True
If indrpt = 1 Then
    indrpt = 0
ElseIf mnthrpt = 1 Then
    mnthrpt = 0
End If
End Sub

Private Sub cmdcurrent_Click()
'current leave status of all employees  report show
emp_leave_status.Show
Unload Me
End Sub

Private Sub cmdexit_Click()
Unload Me
frmrptChoice.Show
End Sub

Private Sub cmdind_Click()
'indivdual employees leave details
indrpt = 1
Call getdata
Shape1.Visible = True
cmdokay.Visible = True
cmdCancel.Visible = True
Label1.Visible = True
Label1.AutoSize = True
Label1.Caption = "Select the employee code / name for which report is required"
Combo1.Visible = False
Combo2.Visible = True
cmdmonth.Visible = False
cmdcurrent.Visible = False
cmdexit.Visible = False
End Sub

Private Sub cmdmonth_Click()
'employees selected months leave report
mnthrpt = 1
Shape1.Visible = True
cmdokay.Visible = True
cmdCancel.Visible = True
Label1.Visible = True
Label1.AutoSize = True
Label1.Caption = "Select the month for which the report is required"
Combo1.Visible = True
Combo2.Visible = False
cmdind.Visible = False
cmdcurrent.Visible = False
cmdexit.Visible = False
End Sub

Private Sub cmdokay_Click()
If mnthrpt = 1 Then
    'to check for the open status of the recordset
    If monleave = True Then
        Payroll_environ.rsMon_leave.Open
        monleave = False
    End If
    Payroll_environ.rsMon_leave.Requery (0)
    'to pass value from form to the data report
    Leave.Sections("section4").Controls.Item("datelabel").Caption = Combo1.Text
    Payroll_environ.rsMon_leave.Filter = "month='" & Combo1.Text & "'"
    'to check for existence of records in the recordset
    If Payroll_environ.rsMon_leave.RecordCount < 1 Then
        MsgBox "Employee leave records not found for specified month", vbCritical, "Payroll : Report Generation"
    Else
        Leave.Show
    End If
ElseIf indrpt = 1 Then
    If sumleave = True Then
        Payroll_environ.rsIndividual_emp_leave.Open
        sumleave = False
    End If
    Call extractempcode
    Payroll_environ.rsIndividual_emp_leave.Requery (0)
    
    Payroll_environ.rsIndividual_emp_leave.Filter = "emp_code='" & empind & "'"
    If Payroll_environ.rsIndividual_emp_leave.RecordCount < 1 Then
        MsgBox "Leave records not found for specified employee ", vbCritical, "Payroll : Report Generation"
    Else
        Ind_empleave_summy.Show
    End If
End If
Combo1.Text = ""
Combo2.Text = ""
End Sub


Private Sub Form_Load()
Me.Top = 1500
Me.Left = 2000
mnthrpt = 0
indrpt = 0
End Sub


Public Sub getdata()
Set db = OpenDatabase(App.Path & "\payroll_db.mdb", , False)
Set leav_rpt = db.OpenRecordset("select emp_code, emp_name from emp_personal", dbOpenDynaset)
leav_rpt.MoveFirst
Do While leav_rpt.EOF = False
    Combo2.AddItem leav_rpt.Fields(0) & "    " & leav_rpt.Fields(1)
    leav_rpt.MoveNext
Loop
End Sub

Public Sub extractempcode()
Dim str, str1, lval As String
Dim i As Integer
'-------------extract empcode from combobox text value
    str = Combo2.Text
    For i = 1 To Len(str) Step 1
        str1 = Mid(str, i, 1)
        If str1 = " " Then
            lval = Left(str, i - 1)
            Exit For
        End If
    Next
    With leav_rpt
        If lval = "" Then
            empind = str
        Else
            empind = lval
        End If
    End With
End Sub

