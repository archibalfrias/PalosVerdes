VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmSystemConfig 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System Configuration"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6285
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSystemConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C6B8A4&
      Height          =   1575
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   2895
      Begin VB.CheckBox chkLocked 
         BackColor       =   &H00C6B8A4&
         Caption         =   "Locked when Idle"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   0
         Width           =   1815
      End
      Begin VB.OptionButton optLocked6 
         BackColor       =   &H00C6B8A4&
         Caption         =   "6 Minutes"
         Height          =   375
         Left            =   1560
         TabIndex        =   21
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton optLocked5 
         BackColor       =   &H00C6B8A4&
         Caption         =   "5 Minutes"
         Height          =   375
         Left            =   1560
         TabIndex        =   20
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optLocked4 
         BackColor       =   &H00C6B8A4&
         Caption         =   "4 Minutes"
         Height          =   375
         Left            =   1560
         TabIndex        =   19
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton optLocked3 
         BackColor       =   &H00C6B8A4&
         Caption         =   "3 Minutes"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton optLocked2 
         BackColor       =   &H00C6B8A4&
         Caption         =   "2 Minutes"
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optLocked1 
         BackColor       =   &H00C6B8A4&
         Caption         =   "1 Minute"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C6B8A4&
      Height          =   1575
      Left            =   3240
      TabIndex        =   7
      Top             =   120
      Width           =   2895
      Begin VB.CommandButton cmdWallPaper 
         Caption         =   "..."
         Height          =   255
         Left            =   2520
         MouseIcon       =   "frmSystemConfig.frx":18B02
         MousePointer    =   99  'Custom
         TabIndex        =   25
         ToolTipText     =   "View Background"
         Top             =   0
         Width           =   300
      End
      Begin VB.CheckBox chkSlide 
         BackColor       =   &H00C6B8A4&
         Caption         =   "Enable Background Slide"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   0
         Width           =   2415
      End
      Begin VB.OptionButton optSlide6 
         BackColor       =   &H00C6B8A4&
         Caption         =   "12 Minutes"
         Height          =   375
         Left            =   1560
         TabIndex        =   13
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton optSlide5 
         BackColor       =   &H00C6B8A4&
         Caption         =   "10 Minutes"
         Height          =   375
         Left            =   1560
         TabIndex        =   12
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optSlide4 
         BackColor       =   &H00C6B8A4&
         Caption         =   "8 Minutes"
         Height          =   375
         Left            =   1560
         TabIndex        =   11
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton optSlide3 
         BackColor       =   &H00C6B8A4&
         Caption         =   "6 Minutes"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton optSlide2 
         BackColor       =   &H00C6B8A4&
         Caption         =   "4 Minutes"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optSlide1 
         BackColor       =   &H00C6B8A4&
         Caption         =   "2 Minutes"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C6B8A4&
      Caption         =   "Quotes Changes Every"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   2895
      Begin VB.OptionButton optQuotes6 
         BackColor       =   &H00C6B8A4&
         Caption         =   "6 Minutes"
         Height          =   375
         Left            =   1560
         TabIndex        =   6
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton optQuotes5 
         BackColor       =   &H00C6B8A4&
         Caption         =   "5 Minutes"
         Height          =   375
         Left            =   1560
         TabIndex        =   5
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optQuotes4 
         BackColor       =   &H00C6B8A4&
         Caption         =   "4 Minutes"
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optQuotes3 
         BackColor       =   &H00C6B8A4&
         Caption         =   "3 Minutes"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton optQuotes2 
         BackColor       =   &H00C6B8A4&
         Caption         =   "2 Minutes"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optQuotes1 
         BackColor       =   &H00C6B8A4&
         Caption         =   "1 Minute"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
   End
   Begin lvButton.lvButtons_H cmdSave 
      Height          =   645
      Left            =   3240
      TabIndex        =   23
      Top             =   1890
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1138
      Caption         =   "Save Setting"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   15396057
      Focus           =   0   'False
      cGradient       =   15396057
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      ImgSize         =   32
      cBack           =   16777215
      mPointer        =   99
      mIcon           =   "frmSystemConfig.frx":18E0C
   End
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   645
      Left            =   3240
      TabIndex        =   24
      Top             =   2610
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1138
      Caption         =   "Close"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   15396057
      Focus           =   0   'False
      cGradient       =   15396057
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      ImgSize         =   32
      cBack           =   16777215
      mPointer        =   99
      mIcon           =   "frmSystemConfig.frx":19126
   End
End
Attribute VB_Name = "frmSystemConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strSetting

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()

If chkLocked.Value = 1 Then
    If optLocked1.Value = False And _
    optLocked2.Value = False And _
    optLocked3.Value = False And _
    optLocked4.Value = False And _
    optLocked5.Value = False And _
    optLocked6.Value = False Then MsgBox "Please Select Idle Time!              ", vbCritical, "Error...": Exit Sub
End If
If chkSlide.Value = 1 Then
    If optSlide1.Value = False And _
    optSlide2.Value = False And _
    optSlide3.Value = False And _
    optSlide4.Value = False And _
    optSlide5.Value = False And _
    optSlide6.Value = False Then MsgBox "Please Select Slide Time!              ", vbCritical, "Error...": Exit Sub
End If
If optQuotes1.Value = False And _
optQuotes2.Value = False And _
optQuotes3.Value = False And _
optQuotes4.Value = False And _
optQuotes5.Value = False And _
optQuotes6.Value = False Then MsgBox "Please Select Quotes Change Time!              ", vbCritical, "Error...": Exit Sub

strSetting = CStr(chkLocked.Value) & "/" & _
             CStr(IIf(optLocked1.Value = True, 60 * 1, IIf(optLocked2.Value = True, 60 * 2, IIf(optLocked3.Value = True, 60 * 3, IIf(optLocked4.Value = True, 60 * 4, IIf(optLocked5.Value = True, 60 * 5, IIf(optLocked6.Value = True, 60 * 6, ""))))))) & "}" & _
             CStr(chkSlide.Value) & "/" & _
             CStr(IIf(optSlide1.Value = True, 60 * 2, IIf(optSlide2.Value = True, 60 * 4, IIf(optSlide3.Value = True, 60 * 6, IIf(optSlide4.Value = True, 60 * 8, IIf(optSlide5.Value = True, 60 * 10, IIf(optSlide6.Value = True, 60 * 12, ""))))))) & "}" & _
             CStr(IIf(optQuotes1.Value = True, 60 * 1, IIf(optQuotes2.Value = True, 60 * 2, IIf(optQuotes3.Value = True, 60 * 3, IIf(optQuotes4.Value = True, 60 * 4, IIf(optQuotes5.Value = True, 60 * 5, IIf(optQuotes6.Value = True, 60 * 6, "")))))))

ConnOmega.Execute "UPDATE tbl_Users_Account " & _
                " SET UserSettings = '" & strSetting & "'" & _
                " WHERE (UserName = '" & gbl_UserName & "')"

gbl_LockWhenIdle = chkLocked.Value
gbl_Idle_Time = IIf(optLocked1.Value = True, 60 * 1, IIf(optLocked2.Value = True, 60 * 2, IIf(optLocked3.Value = True, 60 * 3, IIf(optLocked4.Value = True, 60 * 4, IIf(optLocked5.Value = True, 60 * 5, IIf(optLocked6.Value = True, 60 * 6, ""))))))
gbl_Slides_Background = chkSlide.Value
gbl_Slides_Time = IIf(optSlide1.Value = True, 60 * 2, IIf(optSlide2.Value = True, 60 * 4, IIf(optSlide3.Value = True, 60 * 6, IIf(optSlide4.Value = True, 60 * 8, IIf(optSlide5.Value = True, 60 * 10, IIf(optSlide6.Value = True, 60 * 12, ""))))))
gbl_Quotes_Time = IIf(optQuotes1.Value = True, 60 * 1, IIf(optQuotes2.Value = True, 60 * 2, IIf(optQuotes3.Value = True, 60 * 3, IIf(optQuotes4.Value = True, 60 * 4, IIf(optQuotes5.Value = True, 60 * 5, IIf(optQuotes6.Value = True, 60 * 6, ""))))))

MsgBox "System Setting Change!              ", vbInformation, "Setting"

End Sub

Private Sub cmdWallPaper_Click()
If IsLoaded(frmWallPaper) Then Unload Me: frmWallPaper.ZOrder 0 Else Unload Me: frmWallPaper.Show
End Sub

Private Sub Form_Load()
KeyPreview = True
chkLocked.Value = gbl_LockWhenIdle
Select Case CDbl(gbl_Idle_Time) / 60
    Case 1: optLocked1.Value = True
    Case 2: optLocked2.Value = True
    Case 3: optLocked3.Value = True
    Case 4: optLocked4.Value = True
    Case 5: optLocked5.Value = True
    Case 6: optLocked6.Value = True
End Select
chkSlide.Value = gbl_Slides_Background
Select Case CDbl(gbl_Slides_Time) / 60
    Case 2:  optSlide1.Value = True
    Case 4:  optSlide2.Value = True
    Case 6:  optSlide3.Value = True
    Case 8:  optSlide4.Value = True
    Case 10: optSlide5.Value = True
    Case 12: optSlide6.Value = True
End Select
Select Case CDbl(gbl_Quotes_Time) / 60
    Case 1: optQuotes1.Value = True
    Case 2: optQuotes2.Value = True
    Case 3: optQuotes3.Value = True
    Case 4: optQuotes4.Value = True
    Case 5: optQuotes5.Value = True
    Case 6: optQuotes6.Value = True
End Select
End Sub
