VERSION 5.00
Begin VB.Form frmrptChoice 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payroll Management System "
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   2730
   ScaleWidth      =   6015
   Begin VB.CommandButton rptSalary 
      Caption         =   "Employee Salary Reports "
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
      Left            =   3120
      TabIndex        =   7
      Top             =   1560
      Width           =   2655
   End
   Begin VB.CommandButton rptLoan 
      Caption         =   "Loan Details : Reports"
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
      Left            =   3120
      TabIndex        =   6
      Top             =   960
      Width           =   2655
   End
   Begin VB.CommandButton rptLeave 
      Caption         =   "Leave Details : Reports"
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
      Left            =   3120
      TabIndex        =   5
      Top             =   360
      Width           =   2655
   End
   Begin VB.CommandButton rptDepartment 
      Caption         =   "Department Details Report"
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
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   2655
   End
   Begin VB.CommandButton rptEmployee 
      Caption         =   "Employee Details : Reprots"
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
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   2655
   End
   Begin VB.CommandButton rptBranch 
      Caption         =   "Branch Details Report"
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
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   2655
   End
   Begin VB.CommandButton rptCompany 
      Caption         =   "Company Details Report"
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
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2655
   End
   Begin VB.CommandButton cmdExit 
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
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   2160
      Width           =   2655
   End
End
Attribute VB_Name = "frmrptChoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Top = 1500
Me.Left = 2000
End Sub

Private Sub rptBranch_Click(Index As Integer)
brhrpt.Show
End Sub

Private Sub rptCompany_Click()
'Company_details.Show
End Sub

Private Sub rptDepartment_Click()
Dept_stat.Show
End Sub

Private Sub rptEmployee_Click()
'emprpt_selecttype.Show
Unload Me
End Sub

Private Sub rptLeave_Click()
leav_rptfrm.Show
Unload Me
End Sub

Private Sub rptLoan_Click()
loan_rptfrm.Show
Unload Me
End Sub

Private Sub rptSalary_Click()
salreptfrm.Show
End Sub


