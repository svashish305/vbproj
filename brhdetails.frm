VERSION 5.00
Begin VB.Form brhdetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payroll Management System "
   ClientHeight    =   4500
   ClientLeft      =   1470
   ClientTop       =   330
   ClientWidth     =   7920
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   7920
   Begin VB.CommandButton cmdlast 
      Height          =   350
      Left            =   6840
      Picture         =   "brhdetails.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4000
      Width           =   495
   End
   Begin VB.CommandButton cmdnext 
      Height          =   350
      Left            =   6360
      Picture         =   "brhdetails.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4000
      Width           =   495
   End
   Begin VB.CommandButton cmdprev 
      Height          =   350
      Left            =   5880
      Picture         =   "brhdetails.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4000
      Width           =   495
   End
   Begin VB.CommandButton cmdfirst 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   5400
      Picture         =   "brhdetails.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4000
      Width           =   495
   End
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   120
      TabIndex        =   31
      Top             =   3840
      Width           =   4695
      Begin VB.CommandButton cmdcancel 
         Caption         =   "Cancel "
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
         Left            =   3720
         TabIndex        =   17
         Top             =   200
         Width           =   855
      End
      Begin VB.CommandButton cmdexit 
         BackColor       =   &H80000000&
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
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton cmddel 
         BackColor       =   &H80000000&
         Caption         =   "Del "
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
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton cmdedit 
         BackColor       =   &H80000000&
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
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton cmdsave 
         BackColor       =   &H80000000&
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
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton cmdadd 
         BackColor       =   &H80000000&
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
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   200
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2535
      Left            =   120
      TabIndex        =   26
      Top             =   1320
      Width           =   7695
      Begin VB.TextBox txt_pin 
         DataField       =   "pincode"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   960
         TabIndex        =   8
         Text            =   " "
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox txt_web 
         DataField       =   "web"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   4200
         MaxLength       =   25
         TabIndex        =   12
         Text            =   " "
         Top             =   2040
         Width           =   3015
      End
      Begin VB.TextBox txt_addr3 
         DataField       =   "addr3"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   960
         TabIndex        =   7
         Text            =   " "
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox txt_addr2 
         DataField       =   "addr2"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   960
         TabIndex        =   6
         Text            =   " "
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox txt_email 
         DataField       =   "email"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   4200
         TabIndex        =   11
         Text            =   " "
         Top             =   1560
         Width           =   3015
      End
      Begin VB.TextBox txt_fax 
         DataField       =   "fax"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   5400
         TabIndex        =   10
         Text            =   " "
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox txt_tel 
         DataField       =   "tel"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   5400
         TabIndex        =   9
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txt_addr1 
         DataField       =   "addr1"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   960
         MultiLine       =   -1  'True
         TabIndex        =   5
         Text            =   "brhdetails.frx":1108
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label5 
         Caption         =   "Pincode"
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
         TabIndex        =   34
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "WebSite"
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
         Left            =   3240
         TabIndex        =   32
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "E- Mail "
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
         Left            =   3240
         TabIndex        =   30
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label10 
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
         Left            =   4320
         TabIndex        =   29
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label9 
         Caption         =   "Telephone"
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
         Left            =   4320
         TabIndex        =   28
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label7 
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
         TabIndex        =   27
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      ForeColor       =   &H80000007&
      Height          =   1335
      Left            =   120
      TabIndex        =   22
      Top             =   0
      Width           =   7695
      Begin VB.TextBox txt_dateofcomm 
         DataField       =   "dateofcomm"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   6240
         TabIndex        =   4
         Text            =   " "
         Top             =   885
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         DataField       =   "cmp_code"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   4800
         TabIndex        =   3
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox txt_brhname 
         DataField       =   "brh_name"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1080
         MaxLength       =   25
         TabIndex        =   2
         Top             =   885
         Width           =   2775
      End
      Begin VB.TextBox txt_brhcode 
         DataField       =   "brh_code"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Branch Code"
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
         Left            =   120
         TabIndex        =   33
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Date of Commencement"
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
         Left            =   3960
         TabIndex        =   25
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Name"
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
         TabIndex        =   24
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Company Code"
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
         Left            =   3840
         TabIndex        =   23
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "brhdetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim enableflag As Integer
'------add records to the table-------
Private Sub cmdadd_Click()
Dim code, str, num, alp As String
Dim i, K As Integer

Call clear_all(Me)            'procedure in module

cmdadd.Enabled = False
cmddel.Enabled = False
cmdedit.Enabled = False
cmdexit.Enabled = False
cmdfirst.Enabled = False
cmdnext.Enabled = False
cmdprev.Enabled = False
cmdlast.Enabled = False
cmdsave.Enabled = True
cmdCancel.Enabled = True

If enableflag = 2 Then
     Call txtcmb_disable(Me)    'module procedure
     enableflag = 1
End If

' for automatic generation of branch code:user's choice
If brh.EOF = brh.BOF And brh.RecordCount < 1 Then
chkcode:
    code = InputBox("Please enter the branch code (like BRH10):", "Payroll : Branch Code Generation")
    K = Len(code)
    For i = K To 1 Step -1
        str = Mid(code, i, 1)
        If IsNumeric(str) <> True Then
            alp = Mid(code, 1, i)
            num = Mid(code, i + 1)
            Exit For
        End If
    Next i
    If num = "" Then
        MsgBox "Invalid value for branch code, Please enter like 'BRH10'. ", , "Payroll:Branch Details"
        GoTo chkcode
    Else
        txt_brhcode = code
    End If
Else
    brh.MoveLast
    code = brh(0)
    K = Len(code)
    For i = K To 1 Step -1
        str = Mid(code, i, 1)
        If IsNumeric(str) <> True Then
            alp = Mid(code, 1, i)
            num = Mid(code, i + 1)
            Exit For
        End If
    Next i
        txt_brhcode = alp & (CInt(num) + 1)
        txt_brhname.SetFocus
End If
brh.AddNew
End Sub
'----cancel updations-----------
Private Sub cmdCancel_Click()
brh.CancelUpdate
MsgBox "Update cancelled"
Call chk_displayrec
cmdadd.Enabled = True
cmddel.Enabled = True
cmdedit.Enabled = True
cmdexit.Enabled = True
cmdfirst.Enabled = True
cmdnext.Enabled = True
cmdprev.Enabled = True
cmdlast.Enabled = True
cmdsave.Enabled = False
cmdCancel.Enabled = False

If brh.RecordCount < 1 Then
     Call norec_action
End If

If enableflag = 1 Then
     Call txtcmb_disable(Me)    'module procedure
     enableflag = 2
End If
End Sub

'-------delete records as per users choice--------
Private Sub cmddel_Click()
Dim i As Integer
i = MsgBox("Delete this record ? ", vbYesNo, "Payroll")
If i = vbYes Then
     If Combo1.Text = "" Or txt_brhcode.Text = "" Then
          MsgBox "No record found to delete. " & Chr(13) & "To delete, first select a record.", vbCritical, "Payroll : Delete error"
          Exit Sub
     End If
     brh.Delete
     MsgBox "Record Deleted", , "Payroll"
     Call clear_all(Me)            'procedure in module
     If brh.RecordCount < 1 And brh.BOF = brh.EOF Then
          MsgBox "Zero records in the table now.           " & Chr(13) & "If required start entering records now", vbInformation, "Payroll : Data Entry"
          Call norec_action
     Else
          Call chk_displayrec
     End If
End If
End Sub
'--------edit values in the record----------
Private Sub cmdedit_Click()
cmdadd.Enabled = False
cmddel.Enabled = False
cmdedit.Enabled = False
cmdexit.Enabled = False
cmdfirst.Enabled = False
cmdnext.Enabled = False
cmdprev.Enabled = False
cmdlast.Enabled = False
cmdsave.Enabled = True
cmdCancel.Enabled = True
brh.Edit
Call get_brhdata

If enableflag = 2 Then
    Call txtcmb_disable(Me)    'module procedure
    enableflag = 1
End If
End Sub
'-------exit from the form--------
Private Sub cmdexit_Click()
Call menu_disable
Unload brhdetails
End Sub
'--move to the first record-------
Private Sub cmdfirst_Click()
On Error GoTo err_movfirst
brh.MoveFirst
Call get_brhdata
Exit Sub
err_movfirst:
    MsgBox "Zero records in the Branch Details Table", vbInformation, "Payroll : Data Entry Error"
End Sub
'-----move to the last record-----
Private Sub cmdlast_Click()
On Error GoTo err_movlast
brh.MoveLast
Call get_brhdata
Exit Sub
err_movlast:
    MsgBox "Zero records in the Branch Details Table", vbInformation, "Payroll : Data Entry Error"
End Sub
'----move to the next record--------
Private Sub cmdnext_Click()
On Error GoTo err_movnext
brh.MoveNext
If brh.EOF Then
    MsgBox "Current record is the last record", , "Payroll"
    brh.MovePrevious
    Call get_brhdata
    Exit Sub
End If
Call get_brhdata
Exit Sub
err_movnext:
     MsgBox "Zero records in the Branch Details Table", vbInformation, "Payroll : Data Entry Error"
End Sub
'---move to the previous record-----
Private Sub cmdprev_Click()
On Error GoTo err_movprev
brh.MovePrevious
If brh.BOF = True Then
    MsgBox "Current record is the First record", , "Payroll"
    brh.MoveNext
    Call get_brhdata
    Exit Sub
End If
Call get_brhdata
Exit Sub
err_movprev:
     MsgBox "Zero records in the Branch Details Table", vbInformation, "Payroll : Data Entry Error"
End Sub
'------save records to the table-------
Private Sub cmdsave_Click()
If txt_brhname.Text = "" Then
    MsgBox "Branch name cannot be empty." & Chr(13) & "First enter the branch name and then click on Save. ", vbCritical, "Payroll : Data entry error"
    txt_brhname.SetFocus
    Exit Sub
End If
If Combo1.Text = "" Or Combo1.ListIndex = 0 Then
    MsgBox "Company code is either empty or you have selected the heading. " & Chr(13) & "Select a Company code and name.", vbCritical, "Payroll : Data entry error"
    Combo1.Text = ""
    Combo1.SetFocus
    Exit Sub
End If
cmdadd.Enabled = True
cmddel.Enabled = True
cmdedit.Enabled = True
cmdexit.Enabled = True
cmdfirst.Enabled = True
cmdnext.Enabled = True
cmdprev.Enabled = True
cmdlast.Enabled = True
cmdsave.Enabled = False
cmdCancel.Enabled = False
Call put_brhdata
brh.Update

Call chk_displayrec

If enableflag = 1 Then
    Call txtcmb_disable(Me)    'module procedure
    enableflag = 2
End If
End Sub

'---default settings in form load -------
Private Sub Form_Load()
Dim frgkey_status As Integer
Me.Top = 750
Me.Left = 1500
Set db = OpenDatabase(App.Path & "\payroll_db.mdb", , False)
Set brh = db.OpenRecordset("branch", dbOpenDynaset)
frgkey_status = chk_rscount("company")
Set cmp = db.OpenRecordset("company", dbOpenDynaset)
If frgkey_status < 1 Then
    MsgBox "Zero records in the Company Table" & Chr(13) & "First add records to the Company Table and then start adding records to the Branch Table. ", vbCritical, "Payroll :Data entry error"
    
    Call disableall(Me)           'procedure in module
    
    cmdexit.Enabled = True
    Exit Sub
Else
    '====if brh contains 0 records then only add and exit buttons should be enabled========
     If brh.RecordCount < 1 Then
          Call norec_action
     Else
        Call chk_displayrec
     End If
     txt_brhcode.Enabled = False
     cmdsave.Enabled = False
     cmdCancel.Enabled = False
     enableflag = 1
     'set flag value for enabling text boxes
     Call txtcmb_disable(Me)    'module procedure
     enableflag = 2
     Call get_cmpcode               'get company code from company table
End If
Call menu_disable
End Sub
'-------append values to the table--------------
Public Sub get_brhdata()
With brh
    txt_brhcode = .Fields(0)
    Combo1.Text = .Fields(1)
    txt_brhname = .Fields(2)
    txt_addr1 = .Fields(3)
    txt_addr2 = .Fields(4)
    txt_addr3 = .Fields(5)
    txt_pin = .Fields(6)
    txt_tel = .Fields(7)
    txt_fax = .Fields(8)
    txt_email = .Fields(9)
    txt_web = .Fields(10)
    txt_dateofcomm = .Fields(11)
End With
End Sub
'----append values to the fields in the table---------
Public Sub put_brhdata()
Dim str, str1, lval As String
Dim i As Integer
'-------------extract cmpcode from combobox text value
str = Combo1.Text
For i = 1 To Len(str) Step 1
    str1 = Mid(str, i, 1)
    If str1 = " " Then
        lval = Left(str, i - 1)
        Exit For
    End If
Next
With brh
    .Fields(0) = UCase(txt_brhcode)
    If lval = "" Then
        .Fields(1) = UCase(str)
    Else
        .Fields(1) = UCase(lval)
    End If
    .Fields(2) = UCase(txt_brhname)
    .Fields(3) = UCase(txt_addr1)
    .Fields(4) = UCase(txt_addr2)
    .Fields(5) = UCase(txt_addr3)
    .Fields(6) = Val(txt_pin)
    .Fields(7) = UCase(txt_tel)
    .Fields(8) = UCase(txt_fax)
    .Fields(9) = UCase(txt_email)
    .Fields(10) = UCase(txt_web)
    .Fields(11) = Format(txt_dateofcomm.Text, "dd/mm/yy")
End With
End Sub

Public Sub get_cmpcode()
cmp.MoveFirst
Combo1.AddItem "Code       Name"
Do While cmp.EOF = False
    Combo1.AddItem cmp.Fields(0) & "   " & cmp.Fields(1)
    cmp.MoveNext
Loop
End Sub

Private Sub norec_action()
    cmdsave.Enabled = False
    cmdedit.Enabled = False
    cmddel.Enabled = False
    cmdCancel.Enabled = False
    cmdfirst.Enabled = False
    cmdprev.Enabled = False
    cmdnext.Enabled = False
    cmdlast.Enabled = False
End Sub
Private Sub txt_dateofcomm_LostFocus()
If IsDate(txt_dateofcomm.Text) = False Then
    MsgBox "You have entered a wrong value for date." & Chr(13) & "Valid date format :'dd/mm/yy'", vbCritical, "Payroll : Data entry error"
    txt_dateofcomm.Text = ""
    txt_dateofcomm.SetFocus
End If
End Sub

Public Sub chk_displayrec()
    If brh.RecordCount > 0 Or brh.BOF <> brh.EOF Then
        brh.MoveFirst
        Call get_brhdata
    End If
End Sub
