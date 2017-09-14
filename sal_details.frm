VERSION 5.00
Begin VB.Form SAL_DETAILS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payroll Management System "
   ClientHeight    =   6825
   ClientLeft      =   1065
   ClientTop       =   1485
   ClientWidth     =   10245
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3847.25
   ScaleMode       =   0  'User
   ScaleWidth      =   9016.117
   Begin VB.TextBox tnetsal 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6600
      TabIndex        =   32
      Text            =   " "
      Top             =   5640
      Width           =   1815
   End
   Begin VB.TextBox tgross 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   29
      Text            =   " "
      Top             =   5640
      Width           =   1575
   End
   Begin VB.TextBox tgrade 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8280
      TabIndex        =   5
      Text            =   " "
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox tdesg 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5040
      TabIndex        =   4
      Text            =   " "
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox tdept 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Text            =   " "
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox txt_dofissual 
      Height          =   285
      Left            =   8640
      TabIndex        =   2
      Text            =   " "
      Top             =   240
      Width           =   1455
   End
   Begin VB.ComboBox empcombo 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Text            =   " "
      Top             =   240
      Width           =   4935
   End
   Begin VB.CommandButton cmdexit 
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
      Height          =   355
      Left            =   7635
      TabIndex        =   0
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel "
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
      Height          =   355
      Left            =   6420
      TabIndex        =   37
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton cmddel 
      Caption         =   "Delete"
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
      Height          =   355
      Left            =   5205
      TabIndex        =   36
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton cmdedit 
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
      Height          =   355
      Left            =   3990
      TabIndex        =   35
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
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
      Height          =   355
      Left            =   2775
      TabIndex        =   34
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton cmdissue 
      Caption         =   "Issue"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   355
      Left            =   1560
      TabIndex        =   33
      Top             =   6240
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "DEDUCTIONS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   5160
      TabIndex        =   44
      Top             =   1320
      Width           =   4935
      Begin VB.TextBox dedn 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   285
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   77
         Text            =   " "
         Top             =   3720
         Width           =   975
      End
      Begin VB.TextBox t_splded 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   2760
         TabIndex        =   31
         Text            =   " "
         Top             =   3240
         Width           =   855
      End
      Begin VB.TextBox tded5 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3840
         TabIndex        =   28
         Text            =   " "
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox tded4 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3840
         TabIndex        =   27
         Text            =   " "
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox tloan1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   3840
         TabIndex        =   26
         Text            =   " "
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox tloan1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   3840
         TabIndex        =   25
         Text            =   " "
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox tloan1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   3840
         TabIndex        =   24
         Text            =   " "
         Top             =   870
         Width           =   855
      End
      Begin VB.TextBox tins2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3840
         TabIndex        =   23
         Text            =   " "
         Top             =   390
         Width           =   855
      End
      Begin VB.TextBox tins1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   22
         Text            =   " "
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox titax 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   21
         Text            =   " "
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox tptax 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   20
         Text            =   " "
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox tmed_ded 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   19
         Text            =   " "
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox tgpf 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   18
         Text            =   " "
         Top             =   870
         Width           =   855
      End
      Begin VB.TextBox tpf 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   17
         Text            =   " "
         Top             =   390
         Width           =   855
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   4920
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Label Label14 
         Caption         =   "Total"
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
         Left            =   2040
         TabIndex        =   76
         Top             =   3720
         Width           =   495
      End
      Begin VB.Label Label19 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Special Deduction"
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
         Left            =   960
         TabIndex        =   73
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label lbl_oded5 
         AutoSize        =   -1  'True
         Caption         =   "OthDed"
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
         Left            =   2880
         TabIndex        =   67
         Top             =   2760
         Width           =   690
      End
      Begin VB.Label lbl_oded4 
         AutoSize        =   -1  'True
         Caption         =   "Loss of Pay "
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
         Left            =   2475
         TabIndex        =   66
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label lbl_oded3 
         AutoSize        =   -1  'True
         Caption         =   "Loan 3"
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
         Left            =   2970
         TabIndex        =   65
         Top             =   1800
         Width           =   600
      End
      Begin VB.Label lbl_oded2 
         AutoSize        =   -1  'True
         Caption         =   "Loan 2"
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
         Left            =   2970
         TabIndex        =   64
         Top             =   1320
         Width           =   600
      End
      Begin VB.Label lbl_oded1 
         AutoSize        =   -1  'True
         Caption         =   "Loan 1"
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
         Left            =   2970
         TabIndex        =   63
         Top             =   870
         Width           =   600
      End
      Begin VB.Label lbl_ins2 
         AutoSize        =   -1  'True
         Caption         =   "Insurance2"
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
         Left            =   2595
         TabIndex        =   62
         Top             =   390
         Width           =   975
      End
      Begin VB.Label lbl_ins1 
         AutoSize        =   -1  'True
         Caption         =   "Insurance1"
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
         Left            =   120
         TabIndex        =   61
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label lbl_itax 
         AutoSize        =   -1  'True
         Caption         =   "IncomeTax"
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
         Left            =   90
         TabIndex        =   60
         Top             =   2280
         Width           =   1005
      End
      Begin VB.Label lbl_ptax 
         AutoSize        =   -1  'True
         Caption         =   "P.TAX"
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
         Left            =   525
         TabIndex        =   59
         Top             =   1800
         Width           =   570
      End
      Begin VB.Label lbl_medded 
         AutoSize        =   -1  'True
         Caption         =   " Med Ded"
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
         Left            =   210
         TabIndex        =   58
         Top             =   1320
         Width           =   885
      End
      Begin VB.Label lblgpf 
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
         Left            =   600
         TabIndex        =   57
         Top             =   870
         Width           =   495
      End
      Begin VB.Label lblpf 
         AutoSize        =   -1  'True
         Caption         =   "P.F"
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
         Left            =   795
         TabIndex        =   56
         Top             =   390
         Width           =   300
      End
      Begin VB.Label Label21 
         Caption         =   " "
         Height          =   495
         Left            =   1200
         TabIndex        =   79
         Top             =   3360
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "ALLOWANCES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   240
      TabIndex        =   43
      Top             =   1320
      Width           =   4935
      Begin VB.TextBox allw 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   285
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   75
         Text            =   " "
         Top             =   3720
         Width           =   975
      End
      Begin VB.TextBox t_splallw 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2760
         TabIndex        =   30
         Top             =   3240
         Width           =   855
      End
      Begin VB.TextBox tspl 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3600
         TabIndex        =   14
         Text            =   " "
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox tall1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3600
         TabIndex        =   15
         Text            =   " "
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox tall2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3600
         TabIndex        =   16
         Text            =   " "
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox tlta 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3600
         TabIndex        =   13
         Text            =   " "
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox tconv 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3600
         TabIndex        =   12
         Text            =   " "
         Top             =   870
         Width           =   855
      End
      Begin VB.TextBox twash 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   11
         Text            =   " "
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   6570
         TabIndex        =   52
         Text            =   " "
         Top             =   3240
         Width           =   855
      End
      Begin VB.TextBox tmed 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   10
         Text            =   " "
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox tda 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   9
         Text            =   " "
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox tcca 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   8
         Text            =   " "
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox thra 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   7
         Text            =   " "
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox tbasic 
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   " "
         Top             =   390
         Width           =   855
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   5715
         TabIndex        =   46
         Text            =   " "
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   " Conveyance"
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
         Left            =   2190
         TabIndex        =   69
         Top             =   870
         Width           =   1185
      End
      Begin VB.Line Line1 
         X1              =   4920
         X2              =   0
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Label Label13 
         Caption         =   "Total"
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
         Left            =   1920
         TabIndex        =   74
         Top             =   3720
         Width           =   495
      End
      Begin VB.Label Label18 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Special Allowance"
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
         Left            =   960
         TabIndex        =   72
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   " Washing"
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
         Left            =   195
         TabIndex        =   68
         Top             =   2760
         Width           =   840
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   " D.A"
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
         TabIndex        =   55
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   " OTH2"
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
         Left            =   2790
         TabIndex        =   54
         Top             =   2760
         Width           =   585
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   " OTH1"
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
         Left            =   2790
         TabIndex        =   53
         Top             =   2280
         Width           =   585
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   " L.T.A"
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
         Left            =   2865
         TabIndex        =   51
         Top             =   1320
         Width           =   510
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   " Special"
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
         TabIndex        =   50
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   " Medical"
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
         Left            =   270
         TabIndex        =   49
         Top             =   2280
         Width           =   765
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   " C.C.A"
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
         Left            =   495
         TabIndex        =   48
         Top             =   1320
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   " H.R.A"
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
         Left            =   465
         TabIndex        =   47
         Top             =   840
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   " Basic"
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
         Left            =   1800
         TabIndex        =   45
         Top             =   390
         Width           =   555
      End
      Begin VB.Label Label20 
         Caption         =   " "
         Height          =   495
         Left            =   960
         TabIndex        =   78
         Top             =   3360
         Width           =   2655
      End
   End
   Begin VB.Label Label17 
      Caption         =   "Net Salary "
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
      Left            =   5400
      TabIndex        =   71
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label Label16 
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
      Height          =   255
      Left            =   1920
      TabIndex        =   70
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label lbldoi 
      Caption         =   "Issue Date"
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
      Left            =   7440
      TabIndex        =   42
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   " Grade"
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
      Left            =   7560
      TabIndex        =   41
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Designation"
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
      Left            =   3720
      TabIndex        =   40
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lbldept_code 
      Caption         =   "Department "
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
      Left            =   240
      TabIndex        =   39
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblemp_code 
      Caption         =   "Employee Code "
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
      Left            =   240
      TabIndex        =   38
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "SAL_DETAILS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim enableflag As Integer
Dim code As String
Dim dedtotal As Integer

Private Sub cmdCancel_Click()
MsgBox "Cancelling Update", , "Payroll"
sal_det.CancelUpdate
Call cleardata
empcombo.Text = ""
empcombo.Enabled = False
cmdsave.Enabled = False
cmdCancel.Enabled = False
cmdissue.Enabled = True
cmdedit.Enabled = True
cmddel.Enabled = True
cmdExit.Enabled = True
End Sub

Private Sub cmdedit_Click()
empcombo.Enabled = True
sal_det.Edit
End Sub

Private Sub cmdexit_Click()
Call menu_disable
Unload Me
End Sub

Private Sub cmdissue_Click()
Call cleardata
empcombo.Enabled = True
empcombo.Text = ""
sal_det.AddNew
cmdsave.Enabled = True
cmdCancel.Enabled = True
cmdissue.Enabled = False
cmdExit.Enabled = False
End Sub

Private Sub cmdsave_Click()
Call putdata
sal_det.Update
MsgBox "Saved !"
cmdsave.Enabled = False
cmdCancel.Enabled = False
cmdissue.Enabled = True
cmdedit.Enabled = True
cmdExit.Enabled = True
End Sub

'to fetch dept, desg, grade name for employee code
Private Sub empcombo_LostFocus()
Dim str, str1 As String
Dim i As Integer
'-------------extract cmpcode from combobox text value
str = empcombo.Text
For i = 1 To Len(str) Step 1
    str1 = Mid(str, i, 1)
    If str1 = " " Then
        code = Left(str, i - 1)
        Exit For
    End If
Next
Set info = db.OpenRecordset("select dpt.dpt_name, ds.dsg_name, gr.grd_name from emp_personal emp, department dpt, designation ds, grade gr where emp.emp_code='" & code & "'", dbOpenDynaset)
tdept.Text = info.Fields(0)
tdesg.Text = info.Fields(1)
tgrade.Text = info.Fields(2)
txt_dofissual.Text = Format(Date, "dd/mm/yy")
empcombo.Enabled = False
txt_dofissual.SetFocus
End Sub

Private Sub Form_Load()
Me.Top = 750
Me.Left = 1000
Set db = OpenDatabase(App.Path & "\payroll_db.mdb", , False)
Set sal_det = db.OpenRecordset("salary_details", dbOpenDynaset)
Call getempdata
cmdsave.Enabled = False
cmdCancel.Enabled = False
cmdedit.Enabled = False
cmddel.Enabled = False
empcombo.Enabled = False
Call menu_disable
End Sub

Private Sub t_splallw_LostFocus()
If t_splallw.Text = "" Or IsNumeric(t_splallw.Text) = False Then
     t_splallw.Text = 0
End If
allw = Val(tgross.Text) + Val(t_splallw.Text) - dedtotal
t_splded.SetFocus
End Sub


Private Sub t_splded_LostFocus()
If t_splded.Text = "" Or IsNumeric(t_splded.Text) = False Then
     t_splded.Text = 0
End If
dedn.Text = dedtotal + Val(t_splded)
tnetsal.Text = Val(tgross.Text) + Val(t_splallw.Text) - dedtotal - Val(t_splded)
cmdsave.SetFocus
End Sub

Private Sub tbasic_GotFocus()
'Dim b, pf, gpf, med, ptax, itax As Integer
'Dim ins1, ins2, loan1, loan2, loan3, ded4, ded5 As Integer
'
'
'Set sm = db.OpenRecordset("select * from salary_master where salary_master.emp_code='" & code & "'", dbOpenDynaset)
'If sm.RecordCount < 1 Then
'    MsgBox "Currently no records found in the Salary Master", vbInformation, "Payroll"
'Else
'    tbasic.Text = sm.Fields(1)
'    b = Val(tbasic.Text)
'    thra.Text = CInt((sm.Fields(2) / 100) * b)
'    tcca.Text = CInt((sm.Fields(3) / 100) * b)
'    tda.Text = CInt((sm.Fields(4) / 100) * b)
'    tmed.Text = CInt((sm.Fields(5) / 100) * b)
'    twash.Text = CInt((sm.Fields(6) / 100) * b)
'    tconv.Text = CInt((sm.Fields(7) / 100) * b)
'    tlta.Text = CInt((sm.Fields(8) / 100) * b)
'    tall1.Text = CInt((sm.Fields(9) / 100) * b)
'    tall2.Text = CInt((sm.Fields(10) / 100) * b)
'    tpf.Text = CInt((sm.Fields(11) / 100) * b)
'    pf = Val(tpf.Text)
'    tgpf.Text = CInt((sm.Fields(12) / 100) * b)
'    gpf = Val(tgpf.Text)
'    tmed_ded.Text = CInt((sm.Fields(13) / 100) * b)
'    med = Val(tmed_ded.Text)
'    tptax.Text = CInt((sm.Fields(14) / 100) * b)
'    ptax = Val(tptax.Text)
'    tItax.Text = CInt((sm.Fields(15) / 100) * b)
'    itax = Val(tItax.Text)
'    tins1.Text = CInt((sm.Fields(16) / 100) * b)
'    ins1 = Val(tins1.Text)
'    tins2.Text = CInt((sm.Fields(17) / 100) * b)
'    ins2 = Val(tins2.Text)
'    tded4.Text = CInt((sm.Fields(21) / 100) * b)
'    ded4 = Val(tded4.Text)
'    tded5.Text = CInt((sm.Fields(22) / 100) * b)
'    ded5 = Val(tded5.Text)
'End If
'tgross.Text = sm.Fields(23)
'dedtotal = (pf + gpf + med + ptax + itax + ins1 + ins2 + Val(tloan1) + Val(tloan2) + Val(tloan3) + ded4 + ded5)
't_splallw.SetFocus
End Sub
Public Sub getempdata()
Set empsal = db.OpenRecordset("select emp_code, emp_name from emp_personal", dbOpenDynaset)
    empsal.MoveFirst
    Do While empsal.EOF = False
        empcombo.AddItem empsal.Fields(0) & "   " & empsal.Fields(1)
        empsal.MoveNext
    Loop
End Sub

Public Sub putdata()
If tloan1(1).Text = "" Then
    tloan1(1).Text = 0
End If
If tloan1(2).Text = "" Then
    tloan1(2).Text = 0
End If
If tloan1(3).Text = "" Then
    tloan1(3).Text = 0
End If
With sal_det
     .Fields(0) = code
     .Fields(1) = txt_dofissual
     .Fields(2) = tbasic.Text
     .Fields(3) = thra.Text
     .Fields(4) = tcca.Text
     .Fields(5) = tda.Text
     .Fields(6) = tmed.Text
     .Fields(7) = twash.Text
     .Fields(8) = tconv.Text
     .Fields(9) = tlta.Text
     .Fields(10) = tall1.Text
     .Fields(11) = tall2.Text
     .Fields(12) = tpf.Text
     .Fields(13) = tgpf.Text
     .Fields(14) = tmed_ded.Text
     .Fields(15) = tptax.Text
     .Fields(16) = titax.Text
     .Fields(17) = tins1.Text
     .Fields(18) = tins2.Text
     .Fields(19) = tloan1(1).Text
     .Fields(20) = tloan1(2).Text
     .Fields(21) = tloan1(3).Text
     .Fields(22) = tded4.Text
     .Fields(23) = tded5.Text
     .Fields(24) = t_splallw.Text
     .Fields(25) = t_splded.Text
     .Fields(26) = tgross.Text
     .Fields(27) = tnetsal.Text
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

Private Sub txt_dofissual_LostFocus()
Dim b, pf, gpf, med, ptax, itax As Integer
Dim ins1, ins2, loan1, loan2, loan3, ded4, ded5 As Integer
Dim i As Integer
i = 1
Set loan_sal = db.OpenRecordset("select la.amt_paid, la.bal_instal, la.amt_per_instal from Loan_availed la where la.emp_code='" & code & "'", dbOpenDynaset)
If loan_sal.RecordCount < 1 Then
    MsgBox "Employee not availed any Loan.", vbOKOnly, "Payroll"
Else
    Do While loan_sal.EOF = False
        tloan1(i).Text = loan_sal.Fields(2)
        i = i + 1
        loan_sal.MoveNext
    Loop
End If
Set sm = db.OpenRecordset("select * from salary_master where salary_master.emp_code='" & code & "'", dbOpenDynaset)
If sm.RecordCount < 1 Then
    MsgBox "Currently no records found in the Salary Master", vbInformation, "Payroll"
Else
    tbasic.Text = sm.Fields(1)
    b = Val(tbasic.Text)
    thra.Text = CInt((sm.Fields(2) / 100) * b)
    tcca.Text = CInt((sm.Fields(3) / 100) * b)
    tda.Text = CInt((sm.Fields(4) / 100) * b)
    tmed.Text = CInt((sm.Fields(5) / 100) * b)
    twash.Text = CInt((sm.Fields(6) / 100) * b)
    tconv.Text = CInt((sm.Fields(7) / 100) * b)
    tlta.Text = CInt((sm.Fields(8) / 100) * b)
    tall1.Text = CInt((sm.Fields(9) / 100) * b)
    tall2.Text = CInt((sm.Fields(10) / 100) * b)
    tpf.Text = CInt((sm.Fields(11) / 100) * b)
    pf = Val(tpf.Text)
    tgpf.Text = CInt((sm.Fields(12) / 100) * b)
    gpf = Val(tgpf.Text)
    tmed_ded.Text = CInt((sm.Fields(13) / 100) * b)
    med = Val(tmed_ded.Text)
    tptax.Text = CInt((sm.Fields(14) / 100) * b)
    ptax = Val(tptax.Text)
    titax.Text = CInt((sm.Fields(15) / 100) * b)
    itax = Val(titax.Text)
    tins1.Text = CInt((sm.Fields(16) / 100) * b)
    ins1 = Val(tins1.Text)
    tins2.Text = CInt((sm.Fields(17) / 100) * b)
    ins2 = Val(tins2.Text)
    tded4.Text = CInt((sm.Fields(21) / 100) * b)
    ded4 = Val(tded4.Text)
    tded5.Text = CInt((sm.Fields(22) / 100) * b)
    ded5 = Val(tded5.Text)
End If
tgross.Text = sm.Fields(23)
dedtotal = (pf + gpf + med + ptax + itax + ins1 + ins2 + Val(tloan1(1)) + Val(tloan1(2)) + Val(tloan1(3)) + ded4 + ded5)
t_splallw.SetFocus



End Sub
