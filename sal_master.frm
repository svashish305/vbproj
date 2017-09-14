VERSION 5.00
Begin VB.Form sal_master 
   Caption         =   "Salary Master"
   ClientHeight    =   2400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5580
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   5580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   255
      Left            =   3480
      TabIndex        =   16
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton cmddel 
      Caption         =   "Delete"
      Height          =   255
      Left            =   2760
      TabIndex        =   15
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton cmdedit 
      Caption         =   "Edit"
      Height          =   255
      Left            =   2040
      TabIndex        =   14
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      Height          =   255
      Left            =   1320
      TabIndex        =   13
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "Add"
      Height          =   255
      Left            =   600
      TabIndex        =   12
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   3480
      TabIndex        =   11
      Text            =   " "
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   3480
      TabIndex        =   9
      Text            =   " "
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   3480
      TabIndex        =   7
      Text            =   " "
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   960
      TabIndex        =   5
      Text            =   " "
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Text            =   " "
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Text            =   " "
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Grade Code"
      Height          =   255
      Left            =   2520
      TabIndex        =   10
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "C.C.A"
      Height          =   255
      Left            =   2880
      TabIndex        =   8
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "H.R.A"
      Height          =   255
      Left            =   2880
      TabIndex        =   6
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "D.A"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Basic"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Emp Code"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "sal_master"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
