VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Payroll Management System : SwitchBoard"
   ClientHeight    =   3480
   ClientLeft      =   3885
   ClientTop       =   3075
   ClientWidth     =   8220
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.Menu mnComp 
      Caption         =   "Company Master"
      Begin VB.Menu mnm_comp 
         Caption         =   "&Company"
      End
      Begin VB.Menu mnm_branch 
         Caption         =   "&Branch"
      End
      Begin VB.Menu mnm_dept 
         Caption         =   "&Department"
      End
      Begin VB.Menu mnm_desig 
         Caption         =   "De&signation"
      End
      Begin VB.Menu mnm_grade 
         Caption         =   "&Grade"
      End
   End
   Begin VB.Menu mnEmployee 
      Caption         =   "Employee "
      Begin VB.Menu mnmEmp_pers 
         Caption         =   "Employee_pers"
      End
   End
   Begin VB.Menu mnLeave 
      Caption         =   "Leave "
      Begin VB.Menu mnLevMaster 
         Caption         =   "Leave Master"
      End
      Begin VB.Menu mnLevAvailed 
         Caption         =   "Leave Availed"
      End
   End
   Begin VB.Menu mnLoan 
      Caption         =   "Loan"
      Begin VB.Menu mnLnMast 
         Caption         =   "Loan Master"
      End
      Begin VB.Menu mnLnAvailed 
         Caption         =   "Loan Availed"
      End
   End
   Begin VB.Menu mnSalary 
      Caption         =   "Salary"
      Begin VB.Menu mnPayrollGen 
         Caption         =   "Payroll Generation"
      End
   End
   Begin VB.Menu mn_reports 
      Caption         =   "&Reports"
   End
   Begin VB.Menu mnHelp 
      Caption         =   "&Help"
   End
   Begin VB.Menu mn_exit 
      Caption         =   "E&Xit"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
emprpt = True
sumleave = True
monleave = True
salmonth = True
End Sub

Private Sub mn_exit_Click()
End
End Sub

Private Sub mn_reports_Click()
frmrptChoice.Show
End Sub

Private Sub mnHelp_Click()
MsgBox "Help designing still under progress !!!" & Chr(13) & "Can't show help at this time", vbCritical, "Payroll Help"
End Sub

Private Sub mnLevAvailed_Click()
leave_availed.Show
End Sub
Private Sub mnLevMaster_Click()
leave_master.Show
End Sub

Private Sub mnLnAvailed_Click()
loan.Show
End Sub

Private Sub mnLnMast_Click()
ln_master.Show
End Sub

Private Sub mnm_branch_Click()
brhdetails.Show
End Sub

Private Sub mnm_comp_Click()
cmpdetails.Show
End Sub

Private Sub mnm_dept_Click()
deptdetails.Show
End Sub

Private Sub mnm_desig_Click()
desigdetails.Show
End Sub

Private Sub mnm_grade_Click()
grddetails.Show
End Sub

Private Sub mnmEmp_pers_Click()
emp_pers.Show
End Sub

Private Sub mnPayrollGen_Click()
SAL_DETAILS.Show
End Sub
