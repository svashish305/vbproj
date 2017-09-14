VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5700
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   4515
   ScaleWidth      =   5700
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   4080
      Width           =   1095
   End
   Begin VB.ListBox List2 
      Enabled         =   0   'False
      Height          =   2595
      Left            =   1560
      TabIndex        =   1
      Top             =   1320
      Width           =   3975
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Company Name"
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
      Left            =   1560
      TabIndex        =   5
      Top             =   1080
      Width           =   3975
   End
   Begin VB.Label Label2 
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
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Please select the Company code of the Record which you want to Edit :"
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
      TabIndex        =   3
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdOk_Click()
edit_cmp = List1.List(List1.ListIndex)
MsgBox edit_cmp
Unload Me
End Sub

Private Sub Form_Load()
Dim rs_cmpedit As Recordset
'Dim db As Database
'Set db = OpenDatabase(App.Path & "\payroll_db.mdb", dbOpenDynamic)
Set rs_cmpedit = db.OpenRecordset("select C.cmp_code, C.cmp_name from company C", dbOpenDynaset)
rs_cmpedit.MoveFirst
Do While rs_cmpedit.EOF = False
    List1.AddItem rs_cmpedit.Fields(0)
    List2.AddItem rs_cmpedit.Fields(1)
    rs_cmpedit.MoveNext
Loop
End Sub
