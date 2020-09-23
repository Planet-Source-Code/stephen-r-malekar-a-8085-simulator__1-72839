VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm8085_Sim 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "-:8085 simulator:-"
   ClientHeight    =   8610
   ClientLeft      =   1470
   ClientTop       =   3270
   ClientWidth     =   14835
   Icon            =   "frm8085_Sim.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   14835
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      Caption         =   "Input, Output PORT and Editor"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   8415
      Left            =   11640
      TabIndex        =   79
      Top             =   120
      Width           =   3135
      Begin VB.TextBox txtASMCode 
         Enabled         =   0   'False
         Height          =   2655
         Left            =   120
         ScrollBars      =   3  'Both
         TabIndex        =   90
         Top             =   2760
         Width           =   2895
      End
      Begin VB.Timer Tmr1 
         Enabled         =   0   'False
         Interval        =   750
         Left            =   360
         Top             =   6960
      End
      Begin VB.TextBox txtOUT_Val 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   1320
         TabIndex        =   39
         Text            =   "00"
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox txtOUT_Add 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   240
         TabIndex        =   38
         Text            =   "0000"
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txtIN_Val 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   1320
         TabIndex        =   37
         Text            =   "00"
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtIN_Add 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   240
         TabIndex        =   36
         Text            =   "0000"
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label40 
         Caption         =   "Type the ASM Code:-"
         Height          =   255
         Left            =   120
         TabIndex        =   89
         Top             =   2520
         Width           =   2655
      End
      Begin VB.Label Label39 
         Alignment       =   2  'Center
         BackColor       =   &H0000C0C0&
         Caption         =   $"frm8085_Sim.frx":0442
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   1335
         Left            =   120
         TabIndex        =   88
         Top             =   5475
         Width           =   2895
      End
      Begin VB.Label lbl8085 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "8085 Simulator"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1455
         Left            =   120
         TabIndex        =   87
         Top             =   6840
         Width           =   2895
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000001&
         BorderWidth     =   3
         X1              =   240
         X2              =   2640
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000001&
         BorderWidth     =   3
         X1              =   120
         X2              =   2280
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label Label44 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Address"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   85
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label43 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Value"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1320
         TabIndex        =   84
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label42 
         Caption         =   "Output Port and Address :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   255
         Left            =   240
         TabIndex        =   83
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label Label37 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Address"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   82
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Value"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1320
         TabIndex        =   81
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label35 
         Caption         =   "Input Port and Address :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   255
         Left            =   240
         TabIndex        =   80
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdSetMem 
      Caption         =   "A&pply"
      Height          =   375
      Left            =   10080
      TabIndex        =   7
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox txtPrgAddr 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   345
      Left            =   10200
      TabIndex        =   8
      Text            =   "6000"
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Frame Frame5 
      Caption         =   "System Memory Range (RAM)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2775
      Left            =   120
      TabIndex        =   74
      Top             =   5760
      Width           =   11415
      Begin MSFlexGridLib.MSFlexGrid MSFlxGrdSysMem 
         Height          =   2415
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   4260
         _Version        =   393216
         Cols            =   17
         BackColor       =   8421376
         ForeColor       =   65535
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Bookman Old Style"
            Size            =   9
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.TextBox txtSMT 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   360
      Left            =   10200
      TabIndex        =   6
      Text            =   "6FFF"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtSMF 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   360
      Left            =   10200
      TabIndex        =   5
      Text            =   "6100"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      Caption         =   "Interrupts"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   1095
      Left            =   5760
      TabIndex        =   62
      Top             =   4560
      Width           =   5175
      Begin VB.TextBox Text23 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Text            =   "0"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox Text22 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   840
         TabIndex        =   30
         Text            =   "0"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox Text21 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   1560
         TabIndex        =   31
         Text            =   "0"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox Text20 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   2280
         TabIndex        =   32
         Text            =   "0"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox Text19 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   3000
         TabIndex        =   33
         Text            =   "0"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox Text18 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   3720
         TabIndex        =   34
         Text            =   "0"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox Text17 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   4440
         TabIndex        =   35
         Text            =   "0"
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         Caption         =   "SOD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   375
         Left            =   120
         TabIndex        =   69
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         Caption         =   "SID"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   375
         Left            =   840
         TabIndex        =   68
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         Caption         =   "INTR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   375
         Left            =   1500
         TabIndex        =   67
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         Caption         =   "TRAP"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   375
         Left            =   2200
         TabIndex        =   66
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         Caption         =   "R7.5"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   375
         Left            =   3000
         TabIndex        =   65
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Caption         =   "R6.5"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   375
         Left            =   3720
         TabIndex        =   64
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Caption         =   "R5.5"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   375
         Left            =   4440
         TabIndex        =   63
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Flags"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   1095
      Left            =   5760
      TabIndex        =   51
      Top             =   3360
      Width           =   4215
      Begin VB.TextBox txtFlag 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Index           =   0
         Left            =   3600
         TabIndex        =   28
         Text            =   "0"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtFlag 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Index           =   1
         Left            =   3120
         TabIndex        =   27
         Text            =   "0"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtFlag 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Index           =   2
         Left            =   2640
         TabIndex        =   26
         Text            =   "0"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtFlag 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Index           =   3
         Left            =   2160
         TabIndex        =   25
         Text            =   "0"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtFlag 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Index           =   4
         Left            =   1680
         TabIndex        =   24
         Text            =   "0"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtFlag 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Index           =   5
         Left            =   1200
         TabIndex        =   23
         Text            =   "0"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtFlag 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Index           =   6
         Left            =   720
         TabIndex        =   22
         Text            =   "0"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtFlag 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Index           =   7
         Left            =   240
         TabIndex        =   21
         Text            =   "0"
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Left            =   3600
         TabIndex        =   59
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Left            =   3120
         TabIndex        =   58
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Left            =   2640
         TabIndex        =   57
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Left            =   2160
         TabIndex        =   56
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "AC"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Left            =   1680
         TabIndex        =   55
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Left            =   1200
         TabIndex        =   54
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "Z"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Left            =   720
         TabIndex        =   53
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Left            =   240
         TabIndex        =   52
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Registers"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   3135
      Left            =   5760
      TabIndex        =   40
      Top             =   240
      Width           =   4215
      Begin VB.TextBox txtFlags 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   285
         Left            =   3360
         TabIndex        =   20
         Text            =   "00"
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtPC 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   285
         Left            =   3360
         TabIndex        =   19
         Text            =   "00"
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtSP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   285
         Left            =   3360
         TabIndex        =   18
         Text            =   "00"
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtM 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   285
         Left            =   3360
         TabIndex        =   17
         Text            =   "00"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtL 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   285
         Left            =   1320
         TabIndex        =   16
         Top             =   2760
         Width           =   735
      End
      Begin VB.TextBox txtH 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   285
         Left            =   1320
         TabIndex        =   15
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox txtE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   285
         Left            =   1320
         TabIndex        =   14
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox txtD 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   285
         Left            =   1320
         TabIndex        =   13
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtC 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   285
         Left            =   1320
         TabIndex        =   12
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtB 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   285
         Left            =   1320
         TabIndex        =   11
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtA 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   285
         Left            =   1320
         TabIndex        =   10
         Text            =   "00"
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label38 
         Caption         =   "Flags"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   255
         Left            =   2400
         TabIndex        =   86
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label34 
         Caption         =   "PC"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   255
         Left            =   2400
         TabIndex        =   78
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label32 
         Caption         =   "SP"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   255
         Left            =   2400
         TabIndex        =   77
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label28 
         Caption         =   "Mem."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   255
         Left            =   2400
         TabIndex        =   70
         Top             =   600
         Width           =   975
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   2280
         X2              =   2280
         Y1              =   120
         Y2              =   3120
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Value"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3480
         TabIndex        =   61
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Register"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2400
         TabIndex        =   60
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   255
         Left            =   240
         TabIndex        =   50
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "H"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   255
         Left            =   240
         TabIndex        =   49
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "E"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   255
         Left            =   240
         TabIndex        =   48
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   255
         Left            =   240
         TabIndex        =   47
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Accum."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   255
         Left            =   240
         TabIndex        =   44
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Value"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1320
         TabIndex        =   43
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Register"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Assemble 8085 Program (Program Memory Area)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5535
      Begin VB.CommandButton cmdAtOne 
         Caption         =   "A&t Once"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   4920
         Width           =   1575
      End
      Begin VB.CommandButton cmdStep 
         Caption         =   "&Step by Step"
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Top             =   4920
         Width           =   1815
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrdAssemble 
         Height          =   4095
         Left            =   120
         TabIndex        =   41
         Top             =   720
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   7223
         _Version        =   393216
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   14737632
         ForeColor       =   128
         WordWrap        =   -1  'True
         AllowBigSelection=   -1  'True
         ScrollTrack     =   -1  'True
         TextStyleFixed  =   1
         FocusRect       =   2
         GridLinesFixed  =   3
         AllowUserResizing=   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComDlg.CommonDialog ComnDlgOpen 
         Left            =   120
         Top             =   4800
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "&Open....."
         Height          =   375
         Left            =   1800
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdAssemble 
         Caption         =   "&Assemble"
         Height          =   375
         Left            =   3720
         TabIndex        =   4
         Top             =   4920
         Width           =   1695
      End
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      Caption         =   "Start:"
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
      Height          =   195
      Left            =   10200
      TabIndex        =   76
      Top             =   3480
      Width           =   480
   End
   Begin VB.Label Label31 
      Alignment       =   2  'Center
      Caption         =   "Program Memory Range"
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
      Height          =   435
      Left            =   10080
      TabIndex        =   75
      Top             =   3000
      Width           =   1365
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000001&
      BorderWidth     =   3
      X1              =   10080
      X2              =   11520
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000001&
      BorderWidth     =   3
      X1              =   10080
      X2              =   11520
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      Caption         =   "To:"
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
      Height          =   195
      Left            =   10200
      TabIndex        =   73
      Top             =   1560
      Width           =   300
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "From:"
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
      Height          =   195
      Left            =   10200
      TabIndex        =   72
      Top             =   840
      Width           =   480
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      Caption         =   "System Memory Range"
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
      Height          =   435
      Left            =   10080
      TabIndex        =   71
      Top             =   360
      Width           =   1365
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frm8085_Sim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FSO As FileSystemObject
Dim Label As String, Mnemo As String, Op1 As String, _
Op2 As String, Comment As String, OP_CODE As String, _
FN As String, ASM_Flag As Boolean
Dim Flx_Row As Integer, Flx_Col As Integer, Mem_Space As Integer
Dim opcode As String, operand As String, low As Integer, high As Integer
Dim Lower As String, Higher As String
Dim ad As Long, Counter As Integer
Dim da As Integer, s As String, Start_Step As Boolean

Private Sub cmdAssemble_Click()
Start_Step = True
ASM_Flag = True
cmdOpen_Click
ASM_Flag = False
End Sub

Private Sub cmdAtOne_Click()
Dim Byt As Integer
Dim s As String
ResetDisplay
Me.MSFlexGrdAssemble.Col = 5
For Flx_Row = 1 To Me.MSFlexGrdAssemble.Row - 1
    Me.MSFlexGrdAssemble.Row = Flx_Row
    Me.MSFlexGrdAssemble.ColSel = 0
    Me.MSFlexGrdAssemble.Col = 4
    s = Trim(Me.MSFlexGrdAssemble.Text)
    If Len(s) > 0 Then
        If CInt(Me.MSFlexGrdAssemble.Text) > 0 Then
            Byt = CInt(Me.MSFlexGrdAssemble.Text)
            Me.MSFlexGrdAssemble.Col = 1
            opcode = Me.MSFlexGrdAssemble.Text
        End If
    
        Select Case Byt
            Case 1:
            
            Case 2:
                Me.MSFlexGrdAssemble.Row = Flx_Row + 1
                Me.MSFlexGrdAssemble.Col = 1
                If Me.MSFlexGrdAssemble.Text <> " " Then _
                operand = Hex2Dec(Me.MSFlexGrdAssemble.Text, Len(Me.MSFlexGrdAssemble.Text))
            Case 3:
                Me.MSFlexGrdAssemble.Row = Flx_Row + 1
                Me.MSFlexGrdAssemble.Col = 1
                Lower = Me.MSFlexGrdAssemble.Text
                Me.MSFlexGrdAssemble.Row = Flx_Row + 2
                Higher = Me.MSFlexGrdAssemble.Text
                operand = GetValue(Higher, Lower)
        End Select
        Execute
        Display
    End If
Next
Me.cmdAtOne.Enabled = False
Flx_Row = 0
End Sub

Private Sub cmdOpen_Click()
Dim Cn As ADODB.Connection, Rs As ADODB.Recordset
Dim line As String * 80, Ret_Str As String, Rec_Found As Boolean
Dim objFile As File, objTxtStream As TextStream
Dim ad As Long
Dim da As Integer
Dim i As Integer, j  As Integer
Dim s As String, SQL_Str As String
ResetDisplay
Me.cmdStep.Enabled = True
Me.cmdAtOne.Enabled = True
Me.MSFlexGrdAssemble.Clear
Set_FLX_GRD
On Error GoTo ShowErr
Set Cn = New ADODB.Connection
Cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\My_Mnemonics.mdb;Persist Security Info=False"
Set Rs = New ADODB.Recordset

'Get the Program Address
If Len(txtPrgAddr.Text) < 4 Then
    Exit Sub
End If
ad = Hex2Dec(txtPrgAddr.Text, 4)
s = UCase(Hex(ad))
For i = 1 To 4 - Len(s) Step 1
    s = "0" + s
Next i

'FlexGrd Row and Col.
Flx_Row = 1
Flx_Col = 0
If ASM_Flag = False Then
    ComnDlgOpen.ShowOpen
    line = ComnDlgOpen.FileName
    FN = line
End If


If Len(Trim(line)) = 0 Then Exit Sub
Set objFile = FSO.GetFile(FN)
Set objTxtStream = objFile.OpenAsTextStream(ForReading)

Do While Not objTxtStream.AtEndOfStream
    line = objTxtStream.ReadLine
    Fill_MSFlxGrd line
    
    If Len(Mnemo) > 0 Then
        If Me.MSFlexGrdAssemble.Rows <= Flx_Row Then
           Me.MSFlexGrdAssemble.Rows = Flx_Row + 2
        End If

        'Address Column header
        Me.MSFlexGrdAssemble.Row = Flx_Row
        Me.MSFlexGrdAssemble.Col = 0
        Me.MSFlexGrdAssemble.Text = s

        'Set the MSFlexGRDAssemble
        'Label
        Me.MSFlexGrdAssemble.Row = Flx_Row
        Me.MSFlexGrdAssemble.Col = 2
        If Len(Label) > 0 And Label <> " " Then
            Me.MSFlexGrdAssemble.Text = Trim(Label) & ":"
        End If

        'Mnemonic
        Me.MSFlexGrdAssemble.Row = Flx_Row
        Me.MSFlexGrdAssemble.Col = 3
        If Len(Mnemo) > 0 Then
            Me.MSFlexGrdAssemble.Text = Trim(Mnemo)
        End If

        'Op1
        Me.MSFlexGrdAssemble.Row = Flx_Row
        Me.MSFlexGrdAssemble.Col = 3
        If Len(Op1) > 0 Then
            Me.MSFlexGrdAssemble.Text = Me.MSFlexGrdAssemble.Text & " " & Trim(Op1)
        End If

        'Op2
        Me.MSFlexGrdAssemble.Row = Flx_Row
        Me.MSFlexGrdAssemble.Col = 3
        If Len(Op2) > 0 Then
            Me.MSFlexGrdAssemble.Text = Me.MSFlexGrdAssemble.Text & ", " & Trim(Op2)
        End If

        'Find the Op-Code from the Database.
        With Rs
            .CursorType = adOpenStatic
            .LockType = adLockReadOnly
            'This is the 3 byte instruction, thus memory space is 3.
            If UCase(Mnemo) = "JMP" Or _
               UCase(Mnemo) = "JNZ" Or _
               UCase(Mnemo) = "JNC" Or _
               UCase(Mnemo) = "JC" Or _
               UCase(Mnemo) = "JZ" Or _
               UCase(Mnemo) = "JPO" Or _
               UCase(Mnemo) = "JPE" Or _
               UCase(Mnemo) = "JP" Or _
               UCase(Mnemo) = "JM" Or _
               UCase(Mnemo) = "CNZ" Or _
               UCase(Mnemo) = "CZ" Or _
               UCase(Mnemo) = "CALL" Or _
               UCase(Mnemo) = "CNC" Or _
               UCase(Mnemo) = "CC" Or _
               UCase(Mnemo) = "CPO" Or _
               UCase(Mnemo) = "CPE" Or _
               UCase(Mnemo) = "CP" Or _
               UCase(Mnemo) = "CM" Then

               SQL_Str = "SELECT * FROM Mnemo WHERE Mnemonic LIKE '" & Trim(Mnemo) & _
                "' "
               Rs.Open SQL_Str, Cn, , , adCmdText
               Mem_Space = 3
            'This is the 2 byte instruction but no operand.
            ElseIf UCase(Mnemo) = "ANI" Or _
                   UCase(Mnemo) = "ADI" Or _
                   UCase(Mnemo) = "ACI" Or _
                   UCase(Mnemo) = "SUI" Or _
                   UCase(Mnemo) = "SBI" Or _
                   UCase(Mnemo) = "XRI" Or _
                   UCase(Mnemo) = "ORI" Or _
                   UCase(Mnemo) = "OUT" Or _
                   UCase(Mnemo) = "IN" Or _
                   UCase(Mnemo) = "CPI" Then
               SQL_Str = "SELECT * FROM Mnemo WHERE Mnemonic LIKE '" & Trim(Mnemo) & _
                "' "
               Rs.Open SQL_Str, Cn, , , adCmdText
               Op2 = Left(Op1, 2)
               Mem_Space = 2
            ElseIf Len(Mnemo) > 0 And Len(Op1) > 0 And Len(Op2) > 0 Then
                If Len(Op2) >= 4 Then
                    SQL_Str = "SELECT * FROM Mnemo WHERE Mnemonic LIKE '" & Trim(Mnemo) & _
                    "' AND Operand LIKE '" & Trim(Op1) & "' "

                    Rs.Open SQL_Str, Cn, , , adCmdText
                    If Len(Op2) >= 4 Then
                        Op2 = Left(Op2, 4)
                    End If
                    Mem_Space = 3
                    'e.g. LXI H, 0000H
                ElseIf Len(Op2) >= 2 Then
                    Rs.Open "SELECT * FROM Mnemo WHERE Mnemonic LIKE '" & Trim(Mnemo) & _
                    "' AND Operand LIKE '" & Trim(Op1) & "' AND LEN(Op1) > 1", Cn, , , adCmdText
                    If Len(Op2) >= 2 Then
                        Op2 = Left(Op2, 2)
                    End If
                    Mem_Space = 2
                    'e.g. MVI A, 00H
                ElseIf Len(Op2) = 1 Then
                    'e.g. MOV A, B
                    Rs.Open "SELECT * FROM Mnemo WHERE Mnemonic LIKE '" & Mnemo & _
                    "' AND Operand LIKE '" & Op1 & "' AND Op1 LIKE '" & Op2 & "'", Cn, , , adCmdText
                    Mem_Space = 1
                End If
            ElseIf Len(Op1) >= 4 Then
                Rs.Open "SELECT * FROM Mnemo WHERE Mnemonic LIKE '" & Mnemo & _
                "' ", Cn, , , adCmdText
                Op1 = Left(Op1, 4)
                Mem_Space = 3
            ElseIf Len(Op1) >= 1 Then
                Rs.Open "SELECT * FROM Mnemo WHERE Mnemonic LIKE '" & Mnemo & _
                "' AND Operand LIKE '" & Trim(Op1) & "' ", Cn, , , adCmdText
                Mem_Space = 1
            Else
                If Len(Mnemo) > 0 Then _
                Rs.Open "SELECT * FROM Mnemo WHERE Mnemonic LIKE '" & Mnemo & _
                "' ", Cn, , , adCmdText

                Mem_Space = 1
            End If
            If .State = adStateOpen Then
                If .RecordCount > 0 Then
                    Rec_Found = True
                End If
            End If
        If Rec_Found = False Then
            MsgBox "Program Incorrect! (Instruction not Found.)" & vbCrLf & "Unable to fetch the Mnemonics:-> " & vbCrLf & Mnemo & " " & Op1 & ", " & Op2, vbCritical + vbOKOnly, "Error...."
            Exit Sub
        End If
        'Comment
        Me.MSFlexGrdAssemble.Row = Flx_Row
        Me.MSFlexGrdAssemble.Col = 7
        If Len(Comment) > 0 Then
            Me.MSFlexGrdAssemble.Text = Trim(Comment)
        End If

        If Mem_Space >= 3 And Rec_Found = True Then
            'Opcode/ Hex
            'Address Column header
            Me.MSFlexGrdAssemble.Row = Flx_Row
            Me.MSFlexGrdAssemble.Col = 0
            Me.MSFlexGrdAssemble.Text = s
            'Op Code
            Me.MSFlexGrdAssemble.Row = Flx_Row
            Me.MSFlexGrdAssemble.Col = 1
            Me.MSFlexGrdAssemble.CellAlignment = 4
            Me.MSFlexGrdAssemble.Text = .Fields(0)

            'Byte
            Me.MSFlexGrdAssemble.Row = Flx_Row
            Me.MSFlexGrdAssemble.Col = 4
            Me.MSFlexGrdAssemble.Text = .Fields(6)
            'Machine Cycle
            Me.MSFlexGrdAssemble.Row = Flx_Row
            Me.MSFlexGrdAssemble.Col = 5
            Me.MSFlexGrdAssemble.Text = .Fields(4)
            'T-State
            Me.MSFlexGrdAssemble.Row = Flx_Row
            Me.MSFlexGrdAssemble.Col = 6
            Me.MSFlexGrdAssemble.Text = .Fields(5)

            da = Memory(ad)
            Memory(ad) = da
            ad = ad + 1

            If (ad > 65535) Then
                ad = 0
            End If
            s = UCase(Hex(ad))
            For i = 1 To 4 - Len(s) Step 1
                s = "0" + s
            Next i
            Flx_Row = Flx_Row + 1
            If Me.MSFlexGrdAssemble.Rows <= Flx_Row Then Me.MSFlexGrdAssemble.Rows = Flx_Row + 1
            'Address Column header
            Me.MSFlexGrdAssemble.Row = Flx_Row
            Me.MSFlexGrdAssemble.Col = 0
            Me.MSFlexGrdAssemble.Text = s
            'Mem Addr. Lower Order Byte
            If Me.MSFlexGrdAssemble.Rows <= Flx_Row Then Me.MSFlexGrdAssemble.Rows = Flx_Row + 1
            Me.MSFlexGrdAssemble.Row = Flx_Row
            Me.MSFlexGrdAssemble.Col = 1
            Me.MSFlexGrdAssemble.CellAlignment = 4
            If Len(Op1) >= 4 Then
                Me.MSFlexGrdAssemble.Text = Right(Op1, 2)
            Else
                Me.MSFlexGrdAssemble.Text = Right(Op2, 2)
            End If
            da = Memory(ad)
            Memory(ad) = da
            ad = ad + 1
            If (ad > 65535) Then
                ad = 0
            End If
            s = UCase(Hex(ad))
            For i = 1 To 4 - Len(s) Step 1
                s = "0" + s
            Next i

            Flx_Row = Flx_Row + 1
            If Me.MSFlexGrdAssemble.Rows <= Flx_Row Then Me.MSFlexGrdAssemble.Rows = Flx_Row + 1
            'Address Column header
            Me.MSFlexGrdAssemble.Row = Flx_Row
            Me.MSFlexGrdAssemble.Col = 0
            Me.MSFlexGrdAssemble.Text = s
            'Higher Order Byte
            Me.MSFlexGrdAssemble.Row = Flx_Row
            Me.MSFlexGrdAssemble.Col = 1
            Me.MSFlexGrdAssemble.CellAlignment = 4
            If Len(Op1) >= 4 Then
                Me.MSFlexGrdAssemble.Text = Left(Op1, 2)
            Else
                Me.MSFlexGrdAssemble.Text = Left(Op2, 2)
            End If
            Rec_Found = False
        ElseIf Mem_Space >= 2 And Rec_Found = True Then
            'Address Column header
            Me.MSFlexGrdAssemble.Row = Flx_Row
            Me.MSFlexGrdAssemble.Col = 0
            Me.MSFlexGrdAssemble.Text = s
            'Op Code
            Me.MSFlexGrdAssemble.Row = Flx_Row
            Me.MSFlexGrdAssemble.Col = 1
            Me.MSFlexGrdAssemble.CellAlignment = 4
            Me.MSFlexGrdAssemble.Text = .Fields(0)
            'Byte
            Me.MSFlexGrdAssemble.Row = Flx_Row
            Me.MSFlexGrdAssemble.Col = 4
            Me.MSFlexGrdAssemble.Text = .Fields(6)
            'Machine Cycle
            Me.MSFlexGrdAssemble.Row = Flx_Row
            Me.MSFlexGrdAssemble.Col = 5
            Me.MSFlexGrdAssemble.Text = .Fields(4)
            'T-State
            Me.MSFlexGrdAssemble.Row = Flx_Row
            Me.MSFlexGrdAssemble.Col = 6
            Me.MSFlexGrdAssemble.Text = .Fields(5)

            da = Memory(ad)
            Memory(ad) = da
            ad = ad + 1
            If (ad > 65535) Then
                ad = 0
            End If
            s = UCase(Hex(ad))
            For i = 1 To 4 - Len(s) Step 1
                s = "0" + s
            Next i
            Flx_Row = Flx_Row + 1
            If Me.MSFlexGrdAssemble.Rows <= Flx_Row Then Me.MSFlexGrdAssemble.Rows = Flx_Row + 1
            'Address Column header
            Me.MSFlexGrdAssemble.Row = Flx_Row
            Me.MSFlexGrdAssemble.Col = 0
            Me.MSFlexGrdAssemble.Text = s
            'Mem Addr. Lower Order Byte
            If Me.MSFlexGrdAssemble.Rows <= Flx_Row Then Me.MSFlexGrdAssemble.Rows = Flx_Row + 1
            Me.MSFlexGrdAssemble.Row = Flx_Row
            Me.MSFlexGrdAssemble.Col = 1
            Me.MSFlexGrdAssemble.CellAlignment = 4
            Me.MSFlexGrdAssemble.Text = Right(Op2, 2)
            Rec_Found = False
        Else
            If Rec_Found = True Then
                'Address Column header
                Me.MSFlexGrdAssemble.Row = Flx_Row
                Me.MSFlexGrdAssemble.Col = 0
                Me.MSFlexGrdAssemble.Text = s
                'Op Code
                Me.MSFlexGrdAssemble.Row = Flx_Row
                Me.MSFlexGrdAssemble.Col = 1
                Me.MSFlexGrdAssemble.CellAlignment = 4
                Me.MSFlexGrdAssemble.Text = .Fields(0)
                'Byte
                Me.MSFlexGrdAssemble.Row = Flx_Row
                Me.MSFlexGrdAssemble.Col = 4
                Me.MSFlexGrdAssemble.Text = .Fields(6)
                'Machine Cycle
                Me.MSFlexGrdAssemble.Row = Flx_Row
                Me.MSFlexGrdAssemble.Col = 5
                Me.MSFlexGrdAssemble.Text = .Fields(4)
                'T-State
                Me.MSFlexGrdAssemble.Row = Flx_Row
                Me.MSFlexGrdAssemble.Col = 6
                Me.MSFlexGrdAssemble.Text = .Fields(5)
            End If
        End If

        da = Memory(ad)
        Memory(ad) = da
        ad = ad + 1

        If (ad > 65535) Then
            ad = 0
        End If
        s = UCase(Hex(ad))
        For i = 1 To 4 - Len(s) Step 1
            s = "0" + s
        Next i
        Flx_Row = Flx_Row + 1
        'Clear all contents to fill next
        Label = "": Mnemo = "": Op1 = "": Op2 = "": Comment = ""
        If .State = adStateOpen Then
        .Close
        End If
    End With
    End If
Loop

Assemble
txtPC.Text = s
Flx_Row = 0
Set objTxtStream = Nothing
Set objFile = Nothing

Exit Sub

ShowErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical + vbOKOnly, Mnemo

End Sub

Private Sub cmdSetMem_Click()
Dim i As Byte, j As Integer, k As Integer, s As String

Dim R As Integer, C As Integer, Head As Boolean, Body As Boolean
MSFlxGrdSysMem.Rows = 3
R = 0
C = 2
ad = Hex2Dec("00", 4)
For i = 0 To 15
    da = Memory(ad)
    Memory(ad) = da
    s = UCase(Hex(ad))

    For k = 1 To 2 - Len(s) Step 1
        s = "0" + s
    Next k
    Me.MSFlxGrdSysMem.Row = R
    Me.MSFlxGrdSysMem.Col = C - 1
    Me.MSFlxGrdSysMem.CellAlignment = 4
    Me.MSFlxGrdSysMem.Text = s
    Me.MSFlxGrdSysMem.CellFontBold = True
    ad = ad + 1
    C = C + 1
Next
ad = Hex2Dec(txtSMF.Text, 4)
s = UCase(Hex(ad))
R = 1
C = 0
Do While ad <= Hex2Dec(txtSMT.Text, 4)
    C = 0
    If MSFlxGrdSysMem.Rows <= R Then MSFlxGrdSysMem.Rows = R + 2
    da = Memory(ad)
    Memory(ad) = da
    s = UCase(Hex(ad))

    For i = 0 To 16
        For k = 1 To 4 - Len(s) Step 1
            s = "0" + s
        Next k

        If i = 0 Then
            Me.MSFlxGrdSysMem.Row = R
            Me.MSFlxGrdSysMem.Col = 0
            Me.MSFlxGrdSysMem.CellAlignment = 4
            Me.MSFlxGrdSysMem.Text = s
            Me.MSFlxGrdSysMem.CellFontBold = True
        Else
            Me.MSFlxGrdSysMem.Row = R
            Me.MSFlxGrdSysMem.Col = C
            Me.MSFlxGrdSysMem.CellAlignment = 4
            Me.MSFlxGrdSysMem.Text = "00"
        End If
        C = C + 1
    Next
    ad = ad + 16
    R = R + 1
Loop
End Sub

Private Sub cmdStep_Click()
Dim Byt As Integer
Dim s As String
Me.MSFlexGrdAssemble.Col = 5
Me.MSFlexGrdAssemble.AllowBigSelection = True
If Start_Step = True And Flx_Row > 1 Then
Flx_Row = 1
End If
Flx_Row = Flx_Row + 1
If Flx_Row = Me.MSFlexGrdAssemble.Rows - 1 Then
    Flx_Row = 0
    Me.cmdStep.Enabled = False
    Start_Step = False
    Exit Sub
End If
Me.MSFlexGrdAssemble.Col = 0
Me.MSFlexGrdAssemble.ColSel = 4
Me.MSFlexGrdAssemble.SetFocus

If Flx_Row < Me.MSFlexGrdAssemble.Rows - 1 Then
    Start_Step = False
    Me.MSFlexGrdAssemble.Row = Flx_Row
    Me.MSFlexGrdAssemble.Col = 4
    s = Trim(Me.MSFlexGrdAssemble.Text)
    Me.MSFlexGrdAssemble.RowSel = Flx_Row
    If Len(s) > 0 Then
        If CInt(Me.MSFlexGrdAssemble.Text) > 0 Then
            Byt = CInt(Me.MSFlexGrdAssemble.Text)
            Me.MSFlexGrdAssemble.Col = 1
            opcode = Me.MSFlexGrdAssemble.Text
            Me.MSFlexGrdAssemble.Col = 0
            Me.MSFlexGrdAssemble.ColSel = 4
            Me.MSFlexGrdAssemble.SetFocus
        End If

        Select Case Byt
            Case 1:

            Case 2:
                Me.MSFlexGrdAssemble.Row = Flx_Row + 1
                Me.MSFlexGrdAssemble.Col = 1
                If Me.MSFlexGrdAssemble.Text <> " " Then _
                operand = Hex2Dec(Me.MSFlexGrdAssemble.Text, Len(Me.MSFlexGrdAssemble.Text))
                Me.MSFlexGrdAssemble.Col = 0
                Me.MSFlexGrdAssemble.ColSel = 4
                Me.MSFlexGrdAssemble.SetFocus
            Case 3:
                Me.MSFlexGrdAssemble.Row = Flx_Row + 1
                Me.MSFlexGrdAssemble.Col = 1
                Lower = Me.MSFlexGrdAssemble.Text
                Me.MSFlexGrdAssemble.Row = Flx_Row + 2
                Higher = Me.MSFlexGrdAssemble.Text
                operand = GetValue(Higher, Lower)
                Me.MSFlexGrdAssemble.Col = 0
                Me.MSFlexGrdAssemble.ColSel = 4
                Me.MSFlexGrdAssemble.SetFocus
        End Select
        Me.MSFlexGrdAssemble.Col = 0
        Me.MSFlexGrdAssemble.ColSel = 4
        Me.MSFlexGrdAssemble.SetFocus
        Execute
        Display
    End If
Else
    Flx_Row = 0
    ResetDisplay
End If

End Sub

Private Sub Form_Load()
Start_Step = True
Tmr1.Enabled = True
ASM_Flag = False
Set FSO = New FileSystemObject

With ComnDlgOpen
    .Filter = "All text files|*.txt|All files |*.*"
    .CancelError = False
    .InitDir = App.Path
End With
Set_FLX_GRD

cmdSetMem_Click
End Sub

Private Sub Fill_MSFlxGrd(Str As String)
Dim s As String, Length As Integer, i As Integer, k As Integer, j As Integer
Dim Got_Mnemo As Boolean
Length = Len(Trim(Str))
If InStr(1, Str, ":") > 0 Then
    Label = ""
Else
    Label = " "
End If

Mnemo = ""
Op1 = ""
Op2 = ""

Dim tmp_Str As String
For k = 1 To Len(Trim(Str))
    If (Asc(Mid(Str, k, 1)) >= 65 And Asc(Mid(Str, k, 1)) <= 90) Or _
    (Asc(Mid(Str, k, 1)) >= 97 And Asc(Mid(Str, k, 1)) <= 122) Or _
    (Asc(Mid(Str, k, 1)) >= 48 And Asc(Mid(Str, k, 1)) <= 57) Then
    tmp_Str = tmp_Str & Trim(Mid(Str, k, 1))
    End If
Next

'Find each word in the line
For i = 1 To Length
    If Mid(Str, i, 1) = ";" Then
        Comment = Mid(Str, i, Len(Str))
        Exit For
    End If
        
    If Len(Label) = 0 Then
        s = s & Trim(Mid(Str, i, 1))
        If Mid(Str, i, 1) = ":" Then
            Label = Trim(Mid(s, 1, Len(s) - 1))
            s = ""
        End If
    End If
    
    If Len(Label) > 0 And Len(Mnemo) = 0 Then
        If (Asc(Mid(Str, i, 1)) >= 65 And Asc(Mid(Str, i, 1)) <= 90) Or _
        (Asc(Mid(Str, i, 1)) >= 97 And Asc(Mid(Str, i, 1)) <= 122) Then
            s = s & Trim(Mid(Str, i, 1))
        End If
        
        If (Mid(Str, i, 1) = " " Or Mid(Str, i, 1) = vbTab) And Len(s) > 0 Then
            Mnemo = Trim(Mid(s, 1, Len(s)))
            Got_Mnemo = True
            'If there is no comma i.e. it is a single operand
            s = Mid(Str, i, Len(Str))
            
            'We need to Check if there is an operand after Mnemonic
            
            k = InStr(1, Mid(Str, i, Len(Str)), ",")
            j = InStr(1, Mid(Str, i, Len(Str)), ";")
            
            If j < k And j <> 0 Then
                Op1 = Trim(Mid(s, 1, InStr(1, s, ";") - 2))
                s = ""
                Exit For
            End If
            
            If InStr(1, Mid(Str, i, Len(Str)), ",") = 0 Then
                If InStr(1, Mid(Str, i, Len(Str)), ";") > 0 Then
                    If Len(Mid(Str, i, Len(Str))) > 0 Then
                        s = Trim(Mid(Str, InStr(1, Str, ",") + 1, InStr(1, Str, ";") - 1))
                        Op1 = Trim(Mid(s, i, Len(s)))
                        If Len(Op1) >= 1 And Len(Op1) <= 2 Then
                            s = Op1
                            Op1 = ""
                            For k = 1 To Len(s)
                                If (Asc(Mid(s, k, 1)) >= 65 And Asc(Mid(s, k, 1)) <= 90) Or _
                                (Asc(Mid(s, k, 1)) >= 97 And Asc(Mid(s, k, 1)) <= 122) Or _
                                (Asc(Mid(s, k, 1)) >= 48 And Asc(Mid(s, k, 1)) <= 57) Then
                                    Op1 = Op1 & Trim(Mid(s, k, 1))
                                End If
                            Next
                        End If
                        Comment = Trim(Mid(Str, InStr(1, Str, ";"), Len(Str)))
                    End If
                Else
                    Op1 = Trim(Mid(Str, InStr(1, Str, " ") + 1, Len(Str))): s = ""
                End If
                s = ""
                Exit For
            End If
            s = ""
        End If
        'Sometime the instruction is Single Byte/ Implicit and no space is added.
        If Len(s) = Len(tmp_Str) Then
            Mnemo = s
            s = ""
            Exit For
        End If
        

    End If
    
    'If the instruction is Two byte then proceed.
    If Len(Mnemo) > 0 And Len(Op1) = 0 Then
        If Len(Mid(Str, i, Len(Str))) > 0 Then
            If (Asc(Mid(Str, i, 1)) >= 65 And Asc(Mid(Str, i, 1)) <= 90) Or _
                (Asc(Mid(Str, i, 1)) >= 97 And Asc(Mid(Str, i, 1)) <= 122) Or _
                (Asc(Mid(Str, i, 1)) >= 48 And Asc(Mid(Str, i, 1)) <= 57) Then
                    s = s & Trim(Mid(Str, i, 1))
            End If
            If Mid(Str, i, 1) = "," Then
                Op1 = Trim(Mid(s, 1, Len(s)))
                s = ""
            End If
        End If
    End If
    
    'If the instruction is Three byte then proceed.
    If Len(Op1) > 0 And Len(Op2) = 0 Then
        If InStr(1, Mid(Str, i, Len(Str)), ";") > 0 Then
            If Len(Mid(Str, i, Len(Str))) > 0 Then
                s = Trim(Mid(Str, InStr(1, Str, ",") + 1, InStr(1, Str, ";") - 1))
                Op2 = Trim(Mid(s, 1, InStr(1, s, ";") - 1))
                Comment = Trim(Mid(Str, InStr(1, Str, ";"), Len(Str)))
            End If
        Else
            Op2 = Trim(Mid(Str, InStr(1, Str, ",") + 1, Len(Str))): s = ""
        End If
    End If
Next
'Remove unwanted character
s = Mnemo
Mnemo = ""
For k = 1 To Len(s)
If (Asc(Mid(s, k, 1)) >= 65 And Asc(Mid(s, k, 1)) <= 90) Or _
(Asc(Mid(s, k, 1)) >= 97 And Asc(Mid(s, k, 1)) <= 122) Or _
(Asc(Mid(s, k, 1)) >= 48 And Asc(Mid(s, k, 1)) <= 57) Then
Mnemo = Mnemo & Trim(Mid(s, k, 1))
End If
Next

s = Op1
Op1 = ""
For k = 1 To Len(s)
If (Asc(Mid(s, k, 1)) >= 65 And Asc(Mid(s, k, 1)) <= 90) Or _
(Asc(Mid(s, k, 1)) >= 97 And Asc(Mid(s, k, 1)) <= 122) Or _
(Asc(Mid(s, k, 1)) >= 48 And Asc(Mid(s, k, 1)) <= 57) Then
Op1 = Op1 & Trim(Mid(s, k, 1))
End If
Next

s = Op2
Op2 = ""
For k = 1 To Len(s)
If (Asc(Mid(s, k, 1)) >= 65 And Asc(Mid(s, k, 1)) <= 90) Or _
(Asc(Mid(s, k, 1)) >= 97 And Asc(Mid(s, k, 1)) <= 122) Or _
(Asc(Mid(s, k, 1)) >= 48 And Asc(Mid(s, k, 1)) <= 57) Then
Op2 = Op2 & Trim(Mid(s, k, 1))
End If
Next

s = Label & " # " & Mnemo & " # " & Op1 & " # " & Op2 & " # " & Comment

End Sub

Private Sub Assemble()
Dim Flx_Row As Integer, Flx_Col As Integer
Dim i As Integer
Dim Str As String, tmp_Str As String

For Flx_Row = 1 To Me.MSFlexGrdAssemble.Rows - 1
    Me.MSFlexGrdAssemble.Row = Flx_Row
    Me.MSFlexGrdAssemble.Col = 3
    Str = Me.MSFlexGrdAssemble.Text
    
    If InStr(1, UCase(Str), "JMP") Or _
       InStr(1, UCase(Str), "JNZ") Or _
       InStr(1, UCase(Str), "JZ") Or _
       InStr(1, UCase(Str), "JNC") Or _
       InStr(1, UCase(Str), "JC") Or _
       InStr(1, UCase(Str), "JPE") Or _
       InStr(1, UCase(Str), "JPO") Or _
       InStr(1, UCase(Str), "JM") Or _
       InStr(1, UCase(Str), "JP") Then
       
       Str = Mid(Str, InStr(1, Str, " ") + 1, Len(Str)) & ":"
       
       For i = 1 To Me.MSFlexGrdAssemble.Rows
            Me.MSFlexGrdAssemble.Row = i
            Me.MSFlexGrdAssemble.Col = 2
            
            If UCase(Me.MSFlexGrdAssemble.Text) = UCase(Str) Then
                Me.MSFlexGrdAssemble.Row = i
                Me.MSFlexGrdAssemble.Col = 0
                tmp_Str = Me.MSFlexGrdAssemble.Text
                'Lower Order Byte
                Me.MSFlexGrdAssemble.Row = Flx_Row + 1
                Me.MSFlexGrdAssemble.Col = 1
                Me.MSFlexGrdAssemble.Text = Right(tmp_Str, 2)
                'Higher Order Byte
                Me.MSFlexGrdAssemble.Row = Flx_Row + 2
                Me.MSFlexGrdAssemble.Col = 1
                Me.MSFlexGrdAssemble.Text = Left(tmp_Str, 2)
                Exit For
            End If
       Next
    End If
Next
Flx_Row = 0
End Sub

Private Sub Set_FLX_GRD()

Dim Row_Header(8) As String * 20, i As Byte
Row_Header(1) = "Memory Address": Row_Header(2) = "OpCode/ Hex"
Row_Header(3) = "Label": Row_Header(4) = "Mnemonics"
Row_Header(5) = "Bytes": Row_Header(6) = "Machine Cycle"
Row_Header(7) = "T-States": Row_Header(8) = "Comments"
With MSFlexGrdAssemble
    .Cols = 8
    .Rows = 2
    .FixedCols = 0
    .FixedRows = 1
    .Row = 0
    For i = 0 To 7
        .Col = i
        .Text = Row_Header(i + 1)
    Next
End With

End Sub

Public Sub Execute()
    Dim k As Integer
    opcode = Hex2Dec(opcode, Len(opcode))
    Select Case opcode
        Case 0:
        Case 1:
            B = Hex2Dec(Higher, 2)
            C = Hex2Dec(Lower, 2)
            'Call Split(Hex(operand), B, C)
        Case 2:
            'Set the Value of A in B-C pair
            SetValue Me.txtB.Text, Me.txtC.Text, Hex(A)
        Case 3:
            Call Inx(B, C)
        Case 4:
            Call Inr(B)
        Case 5:
            Call Dcr(B)
        Case 6:
            'Call Mov(B, CInt(operand))
            Call Mov(B, CInt(operand))
        Case 7:
            RotateLeft
        Case 9:
            Call Dad(B, C)
        Case 10:
            A = GetValue(Me.txtB.Text, Me.txtC.Text)
        Case 11:
            Call Dcx(B, C)
        Case 12:
            Call Inr(C)
        Case 13:
            Call Dcr(C)
        Case 14:
            'Call Mov(C, CInt(operand))
            Call Mov(C, CInt(operand))
        Case 15:
            RotateRight
        Case 17:
            D = Hex2Dec(Higher, 2)
            E = Hex2Dec(Lower, 2)
            'Call Split(Hex(operand), D, E)
        Case 18:
            'Memory(Merge(D, E)) = A
            SetValue Me.txtD.Text, Me.txtE.Text, Hex(A)
        Case 19:
            Call Inx(D, E)
            'Me.txtM.Text = GetValue(Me.txtD.Text, Me.txtE.Text)
        Case 20:
            Call Inr(D)
        Case 21:
            Call Dcr(D)
        Case 22:
            Call Mov(D, CInt(operand))
        Case 23:
            Call Ral
        Case 25:
            Call Dad(D, E)
        Case 26:
            'A = GetValue(D, E)
            A = GetValue(Me.txtD.Text, Me.txtE.Text)
        Case 27:
            Call Dcx(D, E)
        Case 28:
            Call Inr(E)
        Case 29:
            Call Dcr(E)
        Case 30:
            Call Mov(E, CInt(operand))
        Case 31:
            Call Rar
        Case 33:
            H = Hex2Dec(Higher, 2)
            L = Hex2Dec(Lower, 2)
            'txtM.Text = Replace(Format(Hex2Dec(operand, 2), "@@"), " ", "0")
        Case 34:
            operand = L
            operand = operand + 1
            If operand > 65535 Then
                operand = 1
            End If
            operand = H
        Case 35:
            Call Inx(H, L)
            txtH.Text = Replace(Format(Hex(H), "@@"), " ", "0")
            txtL.Text = Replace(Format(Hex(L), "@@"), " ", "0")
            'Me.txtM.Text = GetValue(Me.txtH.Text, Me.txtL.Text)
        Case 36:
            Call Inr(H)
        Case 37:
            Call Dcr(H)
        Case 38:
            Call Mov(H, CInt(operand))
        Case 39:
            Call Daa
        Case 41:
            Call Dad(H, L)
        Case 42:
            L = operand
            operand = operand + 1
            If operand > 65535 Then
                operand = 0
            End If
            H = operand
        Case 43:
            Call Dcx(H, L)
            txtH.Text = Replace(Format(Hex(H), "@@"), " ", "0")
            txtL.Text = Replace(Format(Hex(L), "@@"), " ", "0")
            Me.txtM.Text = GetValue(Me.txtH.Text, Me.txtL.Text)
        Case 44:
            Call Inr(L)
        Case 45:
            Call Dcr(L)
        Case 46:
            Call Mov(L, CInt(operand))
        Case 47:
            Call Cma
        Case 49:
            SP = operand
        Case 50:
            'operand = A
            SetValue Higher, Lower, Me.txtA.Text
        Case 51:
            SP = SP + 1
            If SP > 65535 Then
                SP = 0
            End If
        Case 52:
            Call Inr(Hex2Dec(GetValue(Me.txtH.Text, Me.txtL.Text), 2))
            If InStr(1, UCase(GetValue(Me.txtH.Text, Me.txtL.Text)), "A") Or _
            InStr(1, UCase(GetValue(Me.txtH.Text, Me.txtL.Text)), "B") Or _
            InStr(1, UCase(GetValue(Me.txtH.Text, Me.txtL.Text)), "C") Or _
            InStr(1, UCase(GetValue(Me.txtH.Text, Me.txtL.Text)), "D") Or _
            InStr(1, UCase(GetValue(Me.txtH.Text, Me.txtL.Text)), "E") Or _
            InStr(1, UCase(GetValue(Me.txtH.Text, Me.txtL.Text)), "F") Then
                Me.txtM.Text = Hex2Dec(GetValue(Me.txtH.Text, Me.txtL.Text), 2) + 1
                SetValue Me.txtH.Text, Me.txtL.Text, Hex(Me.txtM.Text)
            Else
                Me.txtM.Text = CInt(GetValue(Me.txtH.Text, Me.txtL.Text)) + 1
                SetValue Me.txtH.Text, Me.txtL.Text, Me.txtM.Text 'Hex2Dec(Me.txtM.Text, Len(Me.txtM.Text))
            End If
                
        Case 53:
            Call Dcr(Hex2Dec(GetValue(Me.txtH.Text, Me.txtL.Text), 2))
            If InStr(1, UCase(GetValue(Me.txtH.Text, Me.txtL.Text)), "A") Or _
            InStr(1, UCase(GetValue(Me.txtH.Text, Me.txtL.Text)), "B") Or _
            InStr(1, UCase(GetValue(Me.txtH.Text, Me.txtL.Text)), "C") Or _
            InStr(1, UCase(GetValue(Me.txtH.Text, Me.txtL.Text)), "D") Or _
            InStr(1, UCase(GetValue(Me.txtH.Text, Me.txtL.Text)), "E") Or _
            InStr(1, UCase(GetValue(Me.txtH.Text, Me.txtL.Text)), "F") Then
                Me.txtM.Text = Hex2Dec(GetValue(Me.txtH.Text, Me.txtL.Text), 2) - 1
                SetValue Me.txtH.Text, Me.txtL.Text, Hex(Me.txtM.Text)
            Else
                Me.txtM.Text = CInt(GetValue(Me.txtH.Text, Me.txtL.Text)) - 1
                SetValue Me.txtH.Text, Me.txtL.Text, Me.txtM.Text 'Hex2Dec(Me.txtM.Text, Len(Me.txtM.Text))
            End If
            
            'Me.txtM.Text = Hex2Dec(GetValue(Me.txtH.Text, Me.txtL.Text), 2) - 1
            'SetValue Me.txtH.Text, Me.txtL.Text, Hex2Dec(Me.txtM.Text, Len(Me.txtM.Text))
        Case 54:
            'Merge(H, L) = CInt(operand)
            SetValue Me.txtH.Text, Me.txtL.Text, CStr(Hex(operand))
        Case 55:
            Call SetBit(F, 0)
        Case 56:
            A = operand
        Case 57:
            Call DadSP
        Case 58:
            A = Hex2Dec(operand, 2)
            'Or
            'A = GetValue(Higher, Lower)
        Case 59:
            SP = SP - 1
            If SP < 0 Then
                SP = 65535
            End If
        Case 60:
            Call Inr(A)
        Case 61:
            Call Dcr(A)
        Case 62:
            Call Mov(A, CInt(operand))
        Case 63:
            Call ResetBit(F, 0)
        Case 64:
            Call Mov(B, B)
        Case 65:
            Call Mov(B, C)
        Case 66:
            Call Mov(B, D)
        Case 67:
            Call Mov(B, E)
        Case 68:
            Call Mov(B, H)
        Case 69:
            Call Mov(B, L)
        Case 70:
            Call MovMem2Reg(B)
        Case 71:
            Call Mov(B, A)
        Case 72:
            Call Mov(C, B)
        Case 73:
            Call Mov(C, C)
        Case 74:
            Call Mov(C, D)
        Case 75:
            Call Mov(C, E)
        Case 76:
            Call Mov(C, H)
        Case 77:
            Call Mov(C, L)
        Case 78:
            Call MovMem2Reg(C)
        Case 79:
            Call Mov(C, A)
        Case 80:
            Call Mov(D, B)
        Case 81:
            Call Mov(D, C)
        Case 82:
            Call Mov(D, D)
        Case 83:
            Call Mov(D, E)
        Case 84:
            Call Mov(D, H)
        Case 85:
            Call Mov(D, L)
        Case 86:
            Call MovMem2Reg(D)
        Case 87:
            Call Mov(D, A)
        Case 88:
            Call Mov(E, B)
        Case 89:
            Call Mov(E, C)
        Case 90:
            Call Mov(E, D)
        Case 91:
            Call Mov(E, E)
        Case 92:
            Call Mov(E, H)
        Case 93:
            Call Mov(E, L)
        Case 94:
            Call MovMem2Reg(E)
        Case 95:
            Call Mov(E, A)
        Case 96:
            Call Mov(H, B)
        Case 97:
            Call Mov(H, C)
        Case 98:
            Call Mov(H, D)
        Case 99:
            Call Mov(H, E)
        Case 100:
            Call Mov(H, H)
        Case 101:
            Call Mov(H, L)
        Case 102:
            Call MovMem2Reg(H)
        Case 103:
            Call Mov(H, A)
        Case 104:
            Call Mov(L, B)
        Case 105:
            Call Mov(L, C)
        Case 106:
            Call Mov(L, D)
        Case 107:
            Call Mov(L, E)
        Case 108:
            Call Mov(L, H)
        Case 109:
            Call Mov(L, L)
        Case 110:
            Call MovMem2Reg(L)
        Case 111:
            Call Mov(L, A)
        Case 112:
            Call MovReg2Mem(B)
        Case 113:
            Call MovReg2Mem(C)
        Case 114:
            Call MovReg2Mem(D)
        Case 115:
            Call MovReg2Mem(E)
        Case 116:
            Call MovReg2Mem(H)
        Case 117:
            Call MovReg2Mem(L)
        Case 118:
            flag = False
        Case 119:
            Call MovReg2Mem(A)
        Case 120:
            Call Mov(A, B)
        Case 121:
            Call Mov(A, C)
        Case 122:
            Call Mov(A, D)
        Case 123:
            Call Mov(A, E)
        Case 124:
            Call Mov(A, H)
        Case 125:
            Call Mov(A, L)
        Case 126:
            Call MovMem2Reg(A)
        Case 127:
            Call Mov(A, A)
        Case 128:
            Call Add(A, B)
        Case 129:
            Call Add(A, C)
        Case 130:
            Call Add(A, D)
        Case 131:
            Call Add(A, E)
        Case 132:
            Call Add(A, H)
        Case 133:
            Call Add(A, L)
        Case 134:
            'Call Add(A, GetValue(H, L))
            Call Add(A, Hex2Dec(GetValue(Me.txtH.Text, Me.txtL.Text), 2))
        Case 135:
            Call Add(A, A)
        Case 136:
            Call Adc(A, B)
        Case 137:
            Call Adc(A, C)
        Case 138:
            Call Adc(A, D)
        Case 139:
            Call Adc(A, E)
        Case 140:
            Call Adc(A, H)
        Case 141:
            Call Adc(A, L)
        Case 142:
            Call Adc(A, Hex2Dec(GetValue(Me.txtH.Text, Me.txtL.Text), 2))
        Case 143:
            Call Adc(A, A)
        Case 144:
            Call Subs(A, B)
        Case 145:
            Call Subs(A, C)
        Case 146:
            Call Subs(A, D)
        Case 147:
            Call Subs(A, E)
        Case 148:
            Call Subs(A, H)
        Case 149:
            Call Subs(A, L)
        Case 150:
            Call Subs(A, Hex2Dec(GetValue(Me.txtH.Text, Me.txtL.Text), 2))
        Case 151:
            Call Subs(A, A)
        Case 152:
            Call Sbb(B)
        Case 153:
            Call Sbb(C)
        Case 154:
            Call Sbb(D)
        Case 155:
            Call Sbb(E)
        Case 156:
            Call Sbb(H)
        Case 157:
            Call Sbb(L)
        Case 158:
            Call Sbb(Hex2Dec(GetValue(Me.txtH.Text, Me.txtL.Text), 2))
        Case 159:
            Call Sbb(A)
        Case 160:
            Call Ana(B)
        Case 161:
            Call Ana(C)
        Case 162:
            Call Ana(D)
        Case 163:
            Call Ana(E)
        Case 164:
            Call Ana(H)
        Case 165:
            Call Ana(L)
        Case 166:
            Call Ana(Hex2Dec(GetValue(Me.txtH.Text, Me.txtL.Text), 2))
        Case 167:
            Call Ana(A)
        Case 168:
            Call Xra(B)
        Case 169:
            Call Xra(C)
        Case 170:
            Call Xra(D)
        Case 171:
            Call Xra(E)
        Case 172:
            Call Xra(H)
        Case 173:
            Call Xra(L)
        Case 174:
            Call Xra(Hex2Dec(GetValue(Me.txtH.Text, Me.txtL.Text), 2))
        Case 175:
            Call Xra(A)
        Case 176:
            Call Ora(B)
        Case 177:
            Call Ora(C)
        Case 178:
            Call Ora(D)
        Case 179:
            Call Ora(E)
        Case 180:
            Call Ora(H)
        Case 181:
            Call Ora(L)
        Case 182:
            Call Ora(Hex2Dec(GetValue(Me.txtH.Text, Me.txtL.Text), 2))
        Case 183:
            Call Ora(A)
        Case 184:
            Call Cmp(B)
        Case 185:
            Call Cmp(C)
        Case 186:
            Call Cmp(D)
        Case 187:
            Call Cmp(E)
        Case 188:
            Call Cmp(H)
        Case 189:
            Call Cmp(L)
        Case 190:
            Call Cmp(Hex2Dec(GetValue(Me.txtH.Text, Me.txtL.Text), 2))
        Case 191:
            Call Cmp(A)
        Case 192:
            If GetBit(F, Zero) = 0 Then
                Call Ret
            End If
        Case 193:
            Call Pop(B, C)
        Case 194:
            If GetBit(F, Zero) = 0 Then
                Flx_Row = GetRow(Higher, Lower)
                'PC = operand
            End If
        Case 195:
            Flx_Row = GetRow(Higher, Lower)
            'PC = operand
        Case 196:
            If GetBit(F, Zero) = 0 Then
                Call Kall
            End If
        Case 197:
            Call Push(B, C)
        Case 198:
            Call Add(A, CInt(operand))
        Case 200:
            If GetBit(F, Zero) = 1 Then
                Call Ret
            End If
        Case 201:
            Call Ret
        Case 202:
            If GetBit(F, Zero) = 1 Then
                Flx_Row = GetRow(Higher, Lower)
                'PC = operand
            End If
        Case 204:
            If GetBit(F, Zero) = 1 Then
                Call Kall
            End If
        Case 205:
            Call Kall
        Case 206:
            Call Adc(A, CInt(operand))
        Case 207:
            flag = False
        Case 208:
            If GetBit(F, Carry) = 0 Then
                Call Ret
            End If
        Case 209:
            Call Pop(D, E)
        Case 210:
            If GetBit(F, Carry) = 0 Then
                Flx_Row = GetRow(Higher, Lower)
                'PC = operand
            End If
        Case 211:
            'OUT Port
            If operand = Hex2Dec(Me.txtOUT_Add.Text, Len(Me.txtOUT_Add)) Then
                Me.txtOUT_Val.Text = Replace(Format(Hex(A), "@@"), " ", "0")
                'OR
                'Me.txtOUT_Val.Text =Me.txtA.Text
            Else
                MsgBox "Error in locating the Address: " & Hex(operand), vbCritical + vbOKOnly, "Address Error"
                Exit Sub
            End If
        Case 212:
            If GetBit(F, Carry) = 0 Then
                Call Kall
            End If
        Case 213:
            Call Push(D, E)
        Case 214:
            Call Subs(A, CInt(operand))
        Case 216:
            If GetBit(F, Carry) = 1 Then
                Call Ret
            End If
        Case 218:
            If GetBit(F, Carry) = 1 Then
                Flx_Row = GetRow(Higher, Lower)
                'PC = operand
            End If
        Case 219:
            'IN Port
            If operand = Hex2Dec(Me.txtIN_Add.Text, Len(Me.txtIN_Add)) Then
                A = Hex2Dec(Me.txtIN_Val.Text, Len(Me.txtIN_Val.Text))
            Else
                MsgBox "Error in locating the Address: " & Hex(operand), vbCritical + vbOKOnly, "Address Error"
                Exit Sub
            End If
        Case 220:
            If GetBit(F, Carry) = 1 Then
                Call Kall
            End If
        Case 222:
            Call Sbb(CInt(operand))
        Case 223:
        
        Case 224:
            If GetBit(F, Parity) = 0 Then
                Call Ret
            End If
        Case 225:
            Call Pop(H, L)
        Case 226:
            If GetBit(F, Parity) = 0 Then
                Call Kall
            End If
        Case 227:
            Call Xthl
        Case 228:
            If GetBit(F, Parity) = 0 Then
                Call Kall
            End If
        Case 229:
            Call Push(H, L)
        Case 230:
            Call Ana(CInt(operand))
        Case 231:
        
        Case 232:
            If GetBit(F, Parity) = 1 Then
                Call Ret
            End If
        Case 233:
            PC = Hex2Dec(GetValue(Me.txtH.Text, Me.txtL.Text), 2)
        Case 234:
            If GetBit(F, Parity) = 1 Then
                Flx_Row = GetRow(Higher, Lower)
                'PC = operand
            End If
        Case 235:
            Call Swap(H, D)
            Call Swap(L, E)
        Case 236:
            If GetBit(F, Parity) = 1 Then
                Call Kall
            End If
        Case 238:
            Call Xra(CInt(operand))
        Case 239:
            flag = False
        Case 240:
            If GetBit(F, Sign) = 0 Then
                Call Ret
            End If
        Case 241:
            Call Pop(A, F)
        Case 242:
            If GetBit(F, Sign) = 0 Then
                Flx_Row = GetRow(Higher, Lower)
                'PC = operand
            End If
        Case 244:
            If GetBit(F, Sign) = 0 Then
                Call Kall
            End If
        Case 245:
            Call Push(A, F)
        Case 246:
            Call Ora(CInt(operand))
        Case 248:
            If GetBit(F, Sign) = 1 Then
                Call Ret
            End If
        Case 249:
            SP = Hex2Dec(GetValue(Me.txtH.Text, Me.txtL.Text), 2)
        Case 250:
            If GetBit(F, Sign) = 1 Then
                Flx_Row = GetRow(Higher, Lower)
                'PC = operand
            End If
        Case 252:
            If GetBit(F, Sign) = 1 Then
                Call Kall
            End If
        Case 254:
            Call Cmp(CInt(operand))
    End Select
End Sub

Public Sub Display()
    Dim i As Integer
    txtA.Text = Replace(Format(Hex(A), "@@"), " ", "0")
    txtB.Text = Replace(Format(Hex(B), "@@"), " ", "0")
    txtC.Text = Replace(Format(Hex(C), "@@"), " ", "0")
    txtD.Text = Replace(Format(Hex(D), "@@"), " ", "0")
    txtE.Text = Replace(Format(Hex(E), "@@"), " ", "0")
    txtFlags.Text = Replace(Format(Hex(F), "@@"), " ", "0")
    txtH.Text = Replace(Format(Hex(H), "@@"), " ", "0")
    txtL.Text = Replace(Format(Hex(L), "@@"), " ", "0")
    txtSP.Text = Replace(Format(Hex(SP), "@@@@"), " ", "0")
    
    For i = 0 To 7 Step 1
        If GetBit(F, i) = 1 Then
            txtFlag(i).Text = "1"
        Else
            txtFlag(i).Text = "0"
        End If
    Next i
End Sub

Private Sub MSFlxGrdSysMem_Click()
frmVal.Show
End Sub

Private Sub MSFlxGrdSysMem_KeyDown(KeyCode As Integer, Shift As Integer)
frmVal.Show
End Sub

Public Sub ResetDisplay()
    Dim i As Integer
    txtA.Text = "00"
    txtB.Text = "00"
    txtC.Text = "00"
    txtD.Text = "00"
    txtE.Text = "00"
    txtFlags.Text = "00"
    txtH.Text = "00"
    txtL.Text = "00"
    txtM.Text = "00"
    txtSP.Text = "00"
    txtPC.Text = "00"
    A = 0
    B = 0
    C = 0
    D = 0
    E = 0
    F = 0
    H = 0
    L = 0
    SP = 0
    PC = 0
    For i = 0 To 7 Step 1
        txtFlag(i).Text = "0"
    Next i
End Sub

Private Sub tmr1_Timer()
    If flag Then
        lbl8085.Caption = ""
        flag = False
    Else
        lbl8085.Caption = "8085 Simulator"
        Select Case Counter
            Case 0
                lbl8085.ForeColor = vbRed
            Case 1
                lbl8085.ForeColor = vbCyan
            Case 2
                lbl8085.ForeColor = vbYellow
            Case 3
                lbl8085.ForeColor = vbBlue
            Case 4
                lbl8085.ForeColor = vbGreen
            Case 5
                lbl8085.ForeColor = vbMagenta
            Case 6
                lbl8085.ForeColor = vbWhite
            Case Else
                lbl8085.ForeColor = 15
                Counter = -1
        End Select
        Counter = Counter + 1
        flag = True
    End If
End Sub

