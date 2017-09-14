VERSION 5.00
Begin VB.Form cmpdetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payroll  Management System "
   ClientHeight    =   4545
   ClientLeft      =   2220
   ClientTop       =   2265
   ClientWidth     =   7425
   ControlBox      =   0   'False
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   7425
   Begin VB.CommandButton cmdlast 
      Height          =   350
      Left            =   6480
      Picture         =   "cmpdetails.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton cmdprev 
      Height          =   350
      Left            =   5400
      Picture         =   "cmpdetails.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton cmdnext 
      Height          =   350
      Left            =   5880
      Picture         =   "cmpdetails.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton cmdfirst 
      Height          =   350
      Left            =   4920
      Picture         =   "cmpdetails.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Enabled         =   0   'False
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
      TabIndex        =   17
      Top             =   3720
      Width           =   855
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   120
      TabIndex        =   27
      Top             =   3480
      Width           =   4695
      Begin VB.CommandButton cmdexit 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3000
         TabIndex        =   0
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmddel 
         Caption         =   "Del"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2280
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdedit 
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1560
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdsave 
         Caption         =   "Save"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   840
         TabIndex        =   14
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdadd 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   31
      Top             =   0
      Width           =   7215
      Begin VB.TextBox TxtDate 
         Height          =   285
         Left            =   5880
         TabIndex        =   3
         Top             =   360
         Width           =   1050
      End
      Begin VB.TextBox txt_cmpname 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         MaxLength       =   25
         TabIndex        =   2
         Text            =   " "
         Top             =   720
         Width           =   4095
      End
      Begin VB.TextBox txt_cmpcode 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   1
         Text            =   " "
         Top             =   360
         Width           =   1250
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Date of Commencement"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3360
         TabIndex        =   34
         Top             =   360
         Width           =   2265
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Company Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   360
         TabIndex        =   33
         Top             =   720
         Width           =   1485
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   120
      TabIndex        =   22
      Top             =   1320
      Width           =   7215
      Begin VB.TextBox txt_fax 
         Height          =   285
         Left            =   4320
         MaxLength       =   15
         TabIndex        =   10
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txt_tel2 
         Height          =   285
         Left            =   4320
         MaxLength       =   15
         TabIndex        =   9
         Text            =   " "
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txt_web 
         Height          =   285
         Left            =   4320
         MaxLength       =   25
         TabIndex        =   12
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox txt_addr3 
         Height          =   285
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   6
         Text            =   " "
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox txt_addr2 
         Height          =   285
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   5
         Text            =   " "
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txt_pincode 
         Height          =   285
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   7
         Text            =   " "
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txt_email 
         Height          =   285
         Left            =   4320
         MaxLength       =   30
         TabIndex        =   11
         Text            =   " "
         Top             =   1320
         Width           =   2775
      End
      Begin VB.TextBox txt_tel1 
         Height          =   285
         Left            =   4320
         MaxLength       =   15
         TabIndex        =   8
         Text            =   " "
         Top             =   360
         Width           =   1830
      End
      Begin VB.TextBox txt_addr1 
         Height          =   285
         Left            =   1080
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   4
         Text            =   "cmpdetails.frx":1108
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Tel2 #"
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
         Left            =   3480
         TabIndex        =   30
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "Web Site"
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
         Left            =   3480
         TabIndex        =   29
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Pin Code"
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
         Left            =   120
         TabIndex        =   28
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "E-Mail"
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
         Left            =   3480
         TabIndex        =   26
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Fax"
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
         Left            =   3480
         TabIndex        =   25
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Tel1 # "
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
         Left            =   3480
         TabIndex        =   24
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Address"
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
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "cmpdetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim enableflag As Integer
'-------add record to the table---------
Private Sub cmdadd_Click()
Dim code, str, num, alp As String
Dim i, K As Integer

Call clear_all(Me)            'procedure in module

cmdadd.Enabled = False
cmddel.Enabled = False
cmdedit.Enabled = False
cmdExit.Enabled = False
cmdfirst.Enabled = False
cmdnext.Enabled = False
cmdprev.Enabled = False
cmdlast.Enabled = False
cmdsave.Enabled = True
cmdCancel.Enabled = True

' for automatic generation of branch code:user's choice
If cmp.EOF = cmp.BOF And cmp.RecordCount < 1 Then
chkcode:
    code = InputBox("Please enter the company code (like CMP10):", "Payroll : Branch Code Generation")
    K = Len(code)
    If K > 10 Then
        MsgBox "You have entered a value of more than 10 character length" & Chr(13) & "Please enter only less than 10 characters", vbCritical, "Payroll : Data entry error"
        GoTo chkcode
    End If
    For i = K To 1 Step -1
        str = Mid(code, i, 1)
        If IsNumeric(str) <> True Then
            alp = Mid(code, 1, i)
            num = Mid(code, i + 1)
            Exit For
        End If
    Next i
    If num = "" Then
        MsgBox "Invalid value for company code, Please enter like 'BRH10'. ", , "Payroll:Branch Details"
        GoTo chkcode
    Else
        txt_cmpcode = code
    End If
    If enableflag = 2 Then
        Call txtcmb_disable(Me)    'module procedure
        enableflag = 1
    End If
    txt_cmpname.SetFocus
Else
    cmp.MoveLast
    code = cmp(0)
    K = Len(code)
    For i = K To 1 Step -1
        str = Mid(code, i, 1)
        If IsNumeric(str) <> True Then
            alp = Mid(code, 1, i)
            num = Mid(code, i + 1)
            Exit For
        End If
    Next i
    If enableflag = 2 Then
        Call txtcmb_disable(Me)    'module procedure
         enableflag = 1
    End If
    txt_cmpcode = alp & (CInt(num) + 1)
    txt_cmpcode.SetFocus
End If
cmp.AddNew
End Sub

'----cancel updations-----------
Private Sub cmdCancel_Click()
cmp.CancelUpdate
MsgBox "Update cancelled"
Call chk_displayrec
cmdadd.Enabled = True
cmddel.Enabled = True
cmdedit.Enabled = True
cmdExit.Enabled = True
cmdfirst.Enabled = True
cmdnext.Enabled = True
cmdprev.Enabled = True
cmdlast.Enabled = True

cmdsave.Enabled = False
cmdCancel.Enabled = False

cmp.MoveFirst
Call get_cmpdata
cmdadd.SetFocus
If enableflag = 1 Then
    Call txtcmb_disable(Me)        'module procedure
    enableflag = 2
End If
End Sub
'-------delete a record ------------
Private Sub cmddel_Click()
Dim i As Integer
If txt_cmpcode.Text = "" Or txt_cmpname.Text = "" Then
    MsgBox "No record found to delete. " & Chr(13) & "To delete, first select a record.", vbCritical, "Payroll : Delete error"
    Exit Sub
End If
i = MsgBox("Delete this record ? ", vbYesNo, "Payroll:company Details")
If i = vbYes Then
    cmp.Delete
    MsgBox "Record Deleted", , "Payroll"
    Call clear_all(Me)
    If cmp.RecordCount < 1 Then
        MsgBox "Zero records in the table now.           " & Chr(13) & "If required start entering records now", vbInformation, "Payroll : Data Entry"
        Call disableall(Me)
        cmdadd.Enabled = True
        cmdExit.Enabled = True
    Else
        Call chk_displayrec
    End If
End If
End Sub
'------edit record----------
Private Sub cmdedit_Click()
cmdadd.Enabled = False
cmddel.Enabled = False
cmdedit.Enabled = False
cmdExit.Enabled = False
cmdfirst.Enabled = False
cmdnext.Enabled = False
cmdprev.Enabled = False
cmdlast.Enabled = False
cmdsave.Enabled = True
cmdCancel.Enabled = True
If cmp.EOF Then
    cmp.MovePrevious
ElseIf cmp.BOF Then
    cmp.MoveNext
End If
cmp.Edit
Call get_cmpdata
If enableflag = 2 Then
    Call txtcmb_disable(Me)        'module procedure
    enableflag = 1
End If
txt_cmpname.SetFocus
End Sub
'-------closing the form--------
Private Sub cmdexit_Click()
Unload cmpdetails
'Call menu_disable
End Sub
'-----move to the first record-----------
Private Sub cmdfirst_Click()
On Error GoTo err_movfirst
cmp.MoveFirst
get_cmpdata
Exit Sub
err_movfirst:
     MsgBox "Zero records in the Company Details Table", vbInformation, "Payroll : Data Entry Error"
End Sub
'------move to the last record---------
Private Sub cmdlast_Click()
On Error GoTo err_movlast
cmp.MoveLast
get_cmpdata
Exit Sub
err_movlast:
     MsgBox "Zero records in the Company Details Table", vbInformation, "Payroll : Data Entry Error"
End Sub
'--------move to the next record--------
Private Sub cmdnext_Click()
On Error GoTo err_movnext
cmp.MoveNext
If cmp.EOF Then
    MsgBox "Current record is the last record", , "Payroll"
    cmp.MovePrevious
    Call get_cmpdata
    Exit Sub
End If
Call get_cmpdata
Exit Sub
err_movnext:
     MsgBox "Zero records in the Company Details Table", vbInformation, "Payroll : Data Entry Error"
End Sub
'-----------move to the previous record--------
Private Sub cmdprev_Click()
On Error GoTo err_movprev
cmp.MovePrevious
If cmp.BOF Then
    MsgBox "Current record is the First record"
    cmp.MoveNext
    Call get_cmpdata
    Exit Sub
End If
Call get_cmpdata
Exit Sub
err_movprev:
    MsgBox "Zero records in the Company Details Table", vbInformation, "Payroll : Data Entry Error"
End Sub
'--------save record---------------
Private Sub cmdsave_Click()
'Call chk_nullvalue(txt_cmpname)
If txt_cmpname.Text = "" Then
    MsgBox "Company name cannot be null." & Chr(13) & "Enter the company name and then Click on Save.", vbCritical, "Payroll : Data entry error"
    txt_cmpname.SetFocus
    Exit Sub
End If
cmdadd.Enabled = True
cmddel.Enabled = True
cmdedit.Enabled = True
cmdExit.Enabled = True
cmdfirst.Enabled = True
cmdnext.Enabled = True
cmdprev.Enabled = True
cmdlast.Enabled = True

cmdsave.Enabled = False
cmdCancel.Enabled = False
Call put_cmpdata
cmp.Update
If enableflag = 1 Then
    Call txtcmb_disable(Me)        'module procedure
    enableflag = 2
End If
Call chk_displayrec
End Sub
'default settings in form  load
Private Sub Form_Load()
Dim frgkey_status As Integer
Set db = OpenDatabase(App.Path & "\payroll_db.mdb", , False)
Set cmp = db.OpenRecordset("company", dbOpenDynaset)
frgkey_status = chk_rscount("company")
If frgkey_status < 1 Then
    MsgBox "Zero records in the Company Table" & Chr(13) & "Start adding records.", vbInformation, "Payroll :Data entry error"
    enableflag = 1
    Call disableall(Me)         'procedure in module
    enableflag = 2
    cmdadd.Enabled = True
    cmdExit.Enabled = True
    Exit Sub
Else
    cmdCancel.Enabled = False
    cmdsave.Enabled = False
    enableflag = 1
    Call txtcmb_disable(Me)    'module procedure
    enableflag = 2
    Call chk_displayrec
End If
Me.Top = 750
Me.Left = 2200
'Call menu_disable
End Sub
'to extract values from the database
Public Sub get_cmpdata()
txt_cmpcode = cmp(0)
txt_cmpname = cmp(1)
txt_addr1 = cmp(2)
txt_addr2 = cmp(3)
txt_addr3 = cmp(4)
txt_pincode = cmp(5)
txt_tel1 = cmp(6)
txt_tel2 = cmp(7)
txt_fax = cmp(8)
txt_email = cmp(9)
txt_web = cmp(10)
TxtDate.Text = cmp(11)
End Sub
'to append values to the database
Public Sub put_cmpdata()
cmp(0) = UCase(txt_cmpcode)
cmp(1) = UCase(txt_cmpname)
cmp(2) = UCase(txt_addr1)
cmp(3) = UCase(txt_addr2)
cmp(4) = UCase(txt_addr3)
cmp(5) = Val(txt_pincode)
cmp(6) = UCase(txt_tel1)
cmp(7) = UCase(txt_tel2)
cmp(8) = UCase(txt_fax)
cmp(9) = txt_email
cmp(10) = txt_web
cmp(11) = Format(TxtDate.Text, "dd/mm/yy")
End Sub

Private Sub TxtDate_LostFocus()
If IsDate(TxtDate.Text) = False Then
    MsgBox "You have entered a wrong value for date." & Chr(13) & "Valid date format :'dd/mm/yy'", vbCritical, "Payroll : Data entry error"
    TxtDate.Text = ""
    TxtDate.SetFocus
End If
End Sub

Public Sub chk_displayrec()
    If cmp.RecordCount > 0 Or cmp.BOF <> cmp.EOF Then
        cmp.MoveFirst
        Call get_cmpdata
    End If
End Sub


