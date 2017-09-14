VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8370
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   8370
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   3615
      Left            =   240
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   3555
      ScaleWidth      =   6795
      TabIndex        =   4
      Top             =   2760
      Width           =   6855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   " "
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   " "
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Text            =   " "
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim code, str, alpa As String
Dim num As Long
Dim K, i As Integer

'code = Text1.Text
'For i = Len(code) - 1 To 1 Step -1
'    str = mid(code
'
'




'    code = Text1.Text
'    k = Len(code) - 1
'    For i = k To 1 Step -1
'        str = Mid(code, i, 1)
'        If IsNumeric(str) Then
'            alpa = Mid(code, 1, i)
'            num = Left(code, i)
'            Exit For
'        End If
'    Next
'    If IsNull(num) = True Then
'        If IsNull(alpa) = True Or IsNumeric(alpa) = False Then
'        MsgBox "Invalid value for company code, Please enter like 'EST100'. ", vbInformation, "Payroll:Company Details"
'        Exit Sub
'        End If
'    Else
'        Text2.Text = alpa
'        Text3.Text = num
'    End If
End Sub
