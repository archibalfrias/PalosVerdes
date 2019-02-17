VERSION 5.00
Begin VB.Form frmAcctgPostingRange 
   BackColor       =   &H00EAECD9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3750
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAcctgPostingRange.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtTo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      MaxLength       =   100
      TabIndex        =   3
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox txtFrom 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      MaxLength       =   100
      TabIndex        =   2
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton cmdOK 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   240
      Picture         =   "frmAcctgPostingRange.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   1560
   End
   Begin VB.CommandButton cmdCancel 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1920
      Picture         =   "frmAcctgPostingRange.frx":0F3C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1200
      Width           =   1560
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "TO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   630
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "FROM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   270
      Width           =   1095
   End
End
Attribute VB_Name = "frmAcctgPostingRAnge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim s As String
Dim rs As New ADODB.Recordset

Private Sub Form_Load()
KeyPreview = True
Me.Caption = "GL Posting Range"
Me.Top = (Mainform.ScaleHeight - Me.Height) / 4
Me.Left = (Mainform.ScaleWidth - Me.Width) / 2
End Sub
