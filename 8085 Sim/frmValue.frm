VERSION 5.00
Begin VB.Form frmVal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Addr:Value"
   ClientHeight    =   420
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   420
   ScaleWidth      =   2025
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtValue 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   960
      MaxLength       =   2
      TabIndex        =   1
      Text            =   "00"
      Top             =   50
      Width           =   855
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Value :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   45
      Width           =   735
   End
End
Attribute VB_Name = "frmVal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Me.txtValue.Text = frm8085_Sim.MSFlxGrdSysMem.Text
End Sub

Private Sub txtValue_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
    frm8085_Sim.MSFlxGrdSysMem.Text = Me.txtValue.Text
    Unload Me
End If
End Sub
