VERSION 5.00
Begin VB.Form ln_master 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payroll Management System "
   ClientHeight    =   2820
   ClientLeft      =   2085
   ClientTop       =   1995
   ClientWidth     =   7920
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   7920
   Begin VB.TextBox rinterest 
      Height          =   375
      Left            =   3240
      TabIndex        =   13
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton cmdlast 
      Height          =   375
      Left            =   7200
      Picture         =   "ln_master.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton cmdnext 
      Height          =   375
      Left            =   6720
      Picture         =   "ln_master.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton cmdprev 
      Height          =   375
      Left            =   6240
      Picture         =   "ln_master.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton cmdfirst 
      Height          =   375
      Left            =   5760
      Picture         =   "ln_master.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "E&xit"
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
      Left            =   4680
      TabIndex        =   7
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton cmdcancel 
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
      Height          =   300
      Left            =   3840
      TabIndex        =   6
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton cmddel 
      Caption         =   "&Del"
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
      Left            =   3000
      TabIndex        =   5
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton cmdedit 
      Caption         =   "&Edit"
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
      Left            =   2160
      TabIndex        =   4
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Save"
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
      Left            =   1320
      TabIndex        =   3
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "&Add"
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
      Left            =   480
      TabIndex        =   2
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox txt_lname 
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   360
      Width           =   4215
   End
   Begin VB.Frame Frame1 
      Height          =   700
      Left            =   360
      TabIndex        =   14
      Top             =   1800
      Width           =   5295
   End
   Begin VB.Label Label2 
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
      Left            =   1680
      TabIndex        =   12
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Loan Name "
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
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "ln_master"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim enableflag As Integer
Dim flag As Integer
Private Sub cmdadd_Click()
If enableflag = 2 Then
    Call enable
    enableflag = 1
End If
Call cleardata
Call cmdenable
lm.AddNew
End Sub

Private Sub cmdCancel_Click()
lm.CancelUpdate
If enableflag = 1 Then
    Call enable
    enableflag = 2
End If
Call cmdenable
End Sub

Private Sub cmddel_Click()
Dim i As Integer
i = MsgBox("Delete this record ? ", vbQuestion + vbYesNo, "Payroll")
If i = vbYes Then
    If txt_lname.Text = "" Then
        MsgBox "No record found to delete.", vbCritical, "Payroll"
        Exit Sub
    Else
        lm.Delete
        Call cleardata
        MsgBox "Current record deleted", vbInformation, "Payroll"
    End If
End If
End Sub

Private Sub cmdedit_Click()
Call cmdenable
lm.Edit
If enableflag = 2 Then
    Call enable
    enableflag = 1
End If
End Sub

Private Sub cmdexit_Click()
     Unload Me
End Sub

Private Sub cmdfirst_Click()
lm.MoveFirst
Call getdata
MsgBox "First record ", , "Payroll"
End Sub

Private Sub cmdlast_Click()
lm.MoveLast
Call getdata
MsgBox "Last record ", , "Payroll"
End Sub

Private Sub cmdnext_Click()
lm.MoveNext
If lm.EOF Then
    MsgBox "Current record is the last record", , "Payroll"
    lm.MovePrevious
    Call getdata
    Exit Sub
Else
    Call getdata
End If
End Sub

Private Sub cmdprev_Click()
lm.MovePrevious
If lm.BOF Then
    MsgBox "Current record is the first record", , "Payroll"
    lm.MoveNext
    Call getdata
    Exit Sub
Else
    Call getdata
End If
End Sub

Private Sub cmdsave_Click()
Call putdata
lm.Update
If enableflag = 1 Then
    Call enable
    enableflag = 2
End If
Call cmdenable
End Sub
Private Sub Form_Load()
Me.Top = 1500
Me.Left = 2000
Set db = OpenDatabase(App.Path & "\payroll_db.mdb", , False)
Set lm = db.OpenRecordset("Loan_Master", dbOpenDynaset)
If lm.EOF Then
    cmdfirst.Enabled = False
    cmdnext.Enabled = False
    cmdprev.Enabled = False
    cmdlast.Enabled = False
End If
enableflag = 1
If enableflag = 1 Then
    Call enable
    enableflag = 2
End If
cmdsave.Enabled = False
cmdCancel.Enabled = False
End Sub

Public Sub enable()
txt_lname.Enabled = Not txt_lname.Enabled
rinterest.Enabled = Not rinterest.Enabled
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
txt_lname.Text = lm.Fields(0)
rinterest.Text = lm.Fields(1)
End Sub

Public Sub putdata()
    lm.Fields(0) = txt_lname.Text
    lm.Fields(1) = rinterest
End Sub

Public Sub cleardata()
    txt_lname.Text = ""
End Sub



