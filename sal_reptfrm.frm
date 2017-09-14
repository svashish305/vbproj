VERSION 5.00
Begin VB.Form salreptfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payroll Management System "
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   5985
   Begin VB.CommandButton Command3 
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
      Left            =   4680
      TabIndex        =   3
      Top             =   360
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "sal_reptfrm.frx":0000
      Left            =   3240
      List            =   "sal_reptfrm.frx":0028
      TabIndex        =   2
      Top             =   1320
      Width           =   2415
   End
   Begin VB.CommandButton cmdSalsumy 
      Caption         =   "Salary Summary : Monthly"
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
      Left            =   960
      TabIndex        =   1
      Top             =   600
      Width           =   3495
   End
   Begin VB.CommandButton cmdPayslip 
      Caption         =   "&Payslip "
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
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   5775
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
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
         Left            =   3480
         TabIndex        =   7
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton cmdOk 
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
         Left            =   1440
         TabIndex        =   6
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Select the Month for which Payslip is required "
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
         Left            =   720
         TabIndex        =   5
         Top             =   360
         Width           =   2295
      End
   End
End
Attribute VB_Name = "salreptfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Combo1.Enabled = False
cmdOk.Enabled = False
cmdCancel.Enabled = False
cmdSalsumy.Visible = True
cmdPayslip.Visible = True
End Sub

Private Sub cmdOk_Click()
Dim temp1 As String, temp2 As String
Dim tem1 As Date, tem2 As Date
Dim mo As Integer
Combo1.Enabled = False
cmdOk.Enabled = False
cmdCancel.Enabled = False

'to check for the open status of the recordset
    If salmonth = True Then
        Payroll_environ.rsSalary_summary.Open
        salmonth = False
    End If
    Payroll_environ.rsSalary_summary.Requery (0)
    'to pass value from form to the data report
    salary_summ.Sections("Section4").Controls.Item("lblMonth").Caption = Combo1.Text
'    Select Case Combo1.ListIndex + 1
'        Case 1, 3, 5, 7, 8, 10, 12
'            mo = 31
'        Case 2
'            mo = 28
'        Case Else
'            mo = 30
'        End Select
'    temp1 = Combo1.ListIndex + 1 & "/01/" & Year(Date)
'    temp2 = Combo1.ListIndex + 1 & "/" & mo & "/" & Year(Date)
'    tem1 = CDate(temp1)
'    tem2 = CDate(temp2)
'    'temp1 = Format(temp1, "dd/mm/yy")
'    'temp2 = Format(temp2, "dd/mm/yy")
'    'payroll_environ.rsSalary_summary.Filter=
'    Payroll_environ.rsSalary_summary.Filter = "dateofissual  >= #" & tem1 & "# and dateofissual <=#" & tem2 & "#"
    
'to check for existence of records in the recordset
If Payroll_environ.rsSalary_summary.RecordCount < 1 Then
    MsgBox "no records"
Else
    salary_summ.Show
End If
cmdSalsumy.Visible = True
cmdPayslip.Visible = True
End Sub

Private Sub cmdPayslip_Click()
cmdSalsumy.Visible = False
Combo1.Enabled = True
Combo1.Text = ""
cmdOk.Enabled = True
cmdCancel.Enabled = True
End Sub

Private Sub cmdSalsumy_Click()
cmdPayslip.Visible = False
Combo1.Enabled = True
Combo1.Text = ""
cmdOk.Enabled = True
cmdCancel.Enabled = True
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Top = 1500
Me.Left = 2000
Combo1.Enabled = False
cmdOk.Enabled = False
cmdCancel.Enabled = False
End Sub
