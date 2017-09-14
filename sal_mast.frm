VERSION 5.00
Begin VB.Form salary_master 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payroll Management System "
   ClientHeight    =   5280
   ClientLeft      =   690
   ClientTop       =   2385
   ClientWidth     =   10710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   10710
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
      Height          =   375
      Left            =   4200
      TabIndex        =   40
      Top             =   4680
      Width           =   855
   End
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
      Height          =   375
      Left            =   5520
      TabIndex        =   39
      Top             =   4680
      Width           =   855
   End
   Begin VB.TextBox tcca 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox tbasic 
      Height          =   285
      Left            =   3480
      TabIndex        =   0
      Text            =   " "
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox tgross 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7320
      TabIndex        =   22
      Text            =   " "
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox othded3 
      Height          =   285
      Left            =   9480
      TabIndex        =   19
      Text            =   " "
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox othded2 
      Height          =   285
      Left            =   9480
      TabIndex        =   18
      Text            =   " "
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox othded1 
      Height          =   285
      Left            =   9480
      TabIndex        =   17
      Text            =   " "
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox othded5 
      Height          =   285
      Left            =   9480
      TabIndex        =   21
      Text            =   " "
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox othded4 
      Height          =   285
      Left            =   9480
      TabIndex        =   20
      Text            =   " "
      Top             =   3240
      Width           =   975
   End
   Begin VB.TextBox tinsurance2 
      Height          =   285
      Left            =   9480
      TabIndex        =   16
      Text            =   " "
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox tinsurance1 
      Height          =   285
      Left            =   6960
      TabIndex        =   15
      Text            =   " "
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox tItax 
      Height          =   285
      Left            =   6960
      TabIndex        =   14
      Text            =   " "
      Top             =   3240
      Width           =   975
   End
   Begin VB.TextBox tptax 
      Height          =   285
      Left            =   6960
      TabIndex        =   13
      Text            =   " "
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox tmed_ded 
      Height          =   285
      Left            =   6960
      TabIndex        =   12
      Text            =   " "
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox tgpf 
      Height          =   285
      Left            =   6960
      TabIndex        =   11
      Text            =   " "
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox tpf 
      Height          =   285
      Left            =   6960
      TabIndex        =   10
      Text            =   " "
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox oth_all2 
      Height          =   285
      Left            =   3960
      TabIndex        =   9
      Text            =   " "
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox oth_all1 
      Height          =   285
      Left            =   3960
      TabIndex        =   8
      Text            =   " "
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox tconv 
      Height          =   285
      Left            =   3960
      TabIndex        =   6
      Text            =   " "
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox tlta 
      Height          =   285
      Left            =   3960
      TabIndex        =   7
      Text            =   " "
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox twash 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Text            =   " "
      Top             =   3240
      Width           =   975
   End
   Begin VB.TextBox tmed_all 
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Text            =   " "
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox tda 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Text            =   " "
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox thra 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Text            =   " "
      Top             =   1320
      Width           =   975
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   0
      X2              =   14625
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "L.T.A"
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
      Left            =   3360
      TabIndex        =   49
      Top             =   1320
      Width           =   465
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Washing"
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
      Left            =   360
      TabIndex        =   48
      Top             =   3240
      Width           =   795
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Medical "
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
      Left            =   390
      TabIndex        =   47
      Top             =   2760
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "D.A"
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
      Left            =   825
      TabIndex        =   46
      Top             =   2280
      Width           =   330
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "H.R.A"
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
      TabIndex        =   45
      Top             =   1320
      Width           =   525
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Other 2"
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
      Left            =   3195
      TabIndex        =   44
      Top             =   2280
      Width           =   630
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Others 1"
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
      Left            =   3090
      TabIndex        =   43
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Conveyance "
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
      Left            =   2640
      TabIndex        =   42
      Top             =   2760
      Width           =   1185
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "C.C.A"
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
      Left            =   660
      TabIndex        =   41
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Basic Pay "
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
      Left            =   2400
      TabIndex        =   38
      Top             =   240
      Width           =   960
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "Gross Salary "
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
      Left            =   6000
      TabIndex        =   37
      Top             =   240
      Width           =   1215
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   -3945
      X2              =   10680
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "DEDUCTIONS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   7800
      TabIndex        =   36
      Top             =   760
      Width           =   1335
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      Caption         =   "ALLOWANCES "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   2040
      TabIndex        =   35
      Top             =   760
      Width           =   1425
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "Other 3"
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
      Left            =   8715
      TabIndex        =   34
      Top             =   2760
      Width           =   630
   End
   Begin VB.Label Label24 
      Caption         =   "Other 2"
      Height          =   255
      Left            =   8730
      TabIndex        =   33
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "Other 1"
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
      Left            =   8715
      TabIndex        =   32
      Top             =   1800
      Width           =   630
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "Other 5"
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
      Left            =   8715
      TabIndex        =   31
      Top             =   3720
      Width           =   630
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "Other 4"
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
      Left            =   8715
      TabIndex        =   30
      Top             =   3240
      Width           =   630
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "Insurance 2 "
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
      Left            =   8280
      TabIndex        =   29
      Top             =   1320
      Width           =   1065
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "Insurance 1"
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
      Left            =   5730
      TabIndex        =   28
      Top             =   3720
      Width           =   1020
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "Income Tax"
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
      Left            =   5700
      TabIndex        =   27
      Top             =   3240
      Width           =   1050
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "Prof.Tax"
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
      Left            =   6000
      TabIndex        =   26
      Top             =   2760
      Width           =   750
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Medical "
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
      Left            =   5985
      TabIndex        =   25
      Top             =   2280
      Width           =   765
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "G.P.F"
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
      Left            =   6255
      TabIndex        =   24
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Provident Fund"
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
      Left            =   5400
      TabIndex        =   23
      Top             =   1320
      Width           =   1350
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   5280
      X2              =   5280
      Y1              =   1080
      Y2              =   4440
   End
End
Attribute VB_Name = "salary_master"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub getdata()
With sal_mast
    tbasic.Text = .Fields(1) & ""
    thra.Text = .Fields(2) & ""
    tcca.Text = .Fields(3) & ""
    tda.Text = .Fields(4) & ""
    tmed_all.Text = .Fields(5) & ""
    twash.Text = .Fields(6) & ""
    tconv.Text = .Fields(7) & ""
    tlta.Text = .Fields(8) & ""
    oth_all1.Text = .Fields(9) & ""
    oth_all2.Text = .Fields(10) & ""
    tpf.Text = .Fields(11) & ""
    tgpf.Text = .Fields(12) & ""
    tmed_ded.Text = .Fields(13) & ""
    tptax.Text = .Fields(14) & ""
    titax.Text = .Fields(15) & ""
    tinsurance1.Text = .Fields(16) & ""
    tinsurance2.Text = .Fields(17) & ""
    othded1.Text = .Fields(18) & ""
    othded2.Text = .Fields(19) & ""
    othded3.Text = .Fields(20) & ""
    othded4.Text = .Fields(21) & ""
    othded5.Text = .Fields(22) & ""
    tgross.Text = .Fields(23) & ""
End With
End Sub

Public Sub cleardata()
Dim forclear As Control
For Each forclear In Controls
    If TypeOf forclear Is TextBox Then
        forclear.Text = ""
    End If
Next
End Sub

Public Sub enabletxt()
    Call txtcmb_disable(Me)
'    tbasic.Enabled = False
'    thra.Enabled = False
'    tcca.Enabled = False
'    tda.Enabled = False
'    tmed_all.Enabled = False
'    twash.Enabled = False
'    tconv.Enabled = False
'    tlta.Enabled = False
'    oth_all1.Enabled = False
'    oth_all2.Enabled = False
'    tpf.Enabled = False
'    tgpf.Enabled = False
'    tmed_ded.Enabled = False
'    tptax.Enabled = False
'    titax.Enabled = False
'    tinsurance1.Enabled = False
'    tinsurance2.Enabled = False
'    othded1.Enabled = False
'    othded2.Enabled = False
'    othded3.Enabled = False
'    othded4.Enabled = False
'    othded5.Enabled = False
'    tgross.Enabled = False
End Sub

Private Sub cmdexit_Click()
Unload Me
emp_pers.Show
End Sub

Private Sub cmdokay_Click()
Call putdata
salary_master.Hide
emp_pers.Show
End Sub

Public Sub putdata()
With sal_mast
    .Fields(0) = UCase(emp_pers.txt_empcode.Text)
    .Fields(1) = tbasic.Text
    .Fields(2) = thra.Text
    .Fields(3) = tcca.Text
    .Fields(4) = tda.Text
    .Fields(5) = tmed_all.Text
    .Fields(6) = twash.Text
    .Fields(7) = tconv.Text
    .Fields(8) = tlta.Text
    .Fields(9) = Val(oth_all1.Text)
    .Fields(10) = Val(oth_all2.Text)
    .Fields(11) = tpf.Text
    .Fields(12) = tgpf.Text
    .Fields(13) = tmed_ded.Text
    .Fields(14) = tptax.Text
    .Fields(15) = titax.Text
    .Fields(16) = Val(tinsurance1.Text)
    .Fields(17) = Val(tinsurance2.Text)
    .Fields(18) = Val(othded1.Text)
    .Fields(19) = Val(othded2.Text)
    .Fields(20) = Val(othded3.Text)
    .Fields(21) = Val(othded4.Text)
    .Fields(22) = Val(othded5.Text)
    .Fields(23) = Val(tgross.Text)
End With
End Sub

Private Sub Form_Load()
Me.Top = 1500
Me.Left = 1000
End Sub

Private Sub oth_all1_LostFocus()
oth_all1.Text = chec(oth_all1.Text)
End Sub
Private Sub oth_all2_LostFocus()
oth_all2.Text = chec(oth_all2.Text)
End Sub

Private Sub othded1_LostFocus()
othded1.Text = chec(othded1.Text)
End Sub

Private Sub othded2_LostFocus()
othded2.Text = chec(othded2.Text)
End Sub

Private Sub othded3_LostFocus()
othded3.Text = chec(othded3.Text)
End Sub
Private Sub othded4_LostFocus()
othded4.Text = chec(othded4.Text)
End Sub
Private Sub othded5_LostFocus()
othded5.Text = chec(othded5.Text)
End Sub
Public Function chec(txt As String) As Single
If IsNull(txt) Then
    chec = 0
Else
    chec = Val(txt)
End If
tgross.Text = tgross.Text + tbasic.Text * chec / 100
End Function

Private Sub tbasic_LostFocus()
If IsNull(tbasic.Text) = True Then
    MsgBox "Basic salary field is empty. " & Chr(13) & "You have to enter a value in this field as the salary breakage is dependent on basic.", vbCritical, "Payroll : Data entry error"
    tbasic.Text = ""
    tbasic.SetFocus
    Exit Sub
End If
If IsNumeric(tbasic.Text) = False Then
    MsgBox "Basic salary is a numeric value. So, enter a number only", vbCritical, "Payroll : Data entry error"
    tbasic.Text = ""
    tbasic.SetFocus
    Exit Sub
End If
tgross.Text = Val(tbasic.Text)
MsgBox "Allowances and Deductions in Percentage" & Chr(13) & "Please enter the Salary Allowances and Deductions in terms of Percentage only.", vbExclamation, "Payroll"
End Sub

Private Sub tcca_LostFocus()
tcca.Text = chec(tcca.Text)
End Sub

Private Sub tconv_LostFocus()
tconv.Text = chec(tconv.Text)
End Sub

Private Sub tda_LostFocus()
tda.Text = chec(tda.Text)
End Sub

Private Sub tgpf_LostFocus()
tgpf.Text = chec(tgpf.Text)
End Sub

Private Sub thra_LostFocus()
thra.Text = chec(thra.Text)
End Sub

Private Sub tinsurance1_LostFocus()
tinsurance1.Text = chec(tinsurance1.Text)
End Sub
Private Sub tinsurance2_LostFocus()
tinsurance2.Text = chec(tinsurance2.Text)
End Sub

Private Sub tItax_LostFocus()
titax.Text = chec(titax.Text)
End Sub

Private Sub tlta_LostFocus()
tlta.Text = chec(tlta.Text)
End Sub

Private Sub tmed_all_LostFocus()
tmed_all.Text = chec(tmed_all.Text)
End Sub

Private Sub tmed_ded_LostFocus()
tmed_ded.Text = chec(tmed_ded.Text)
End Sub

Private Sub tpf_LostFocus()
tpf.Text = chec(tpf.Text)
End Sub

Private Sub tptax_LostFocus()
tptax.Text = chec(tptax.Text)
End Sub

Private Sub twash_LostFocus()
twash.Text = chec(twash.Text)
End Sub
