VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VB MsgBox Creator"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12030
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   521
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   802
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPreview 
      BackColor       =   &H00FFF8E8&
      Caption         =   "Preview"
      Height          =   255
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   6000
      Width           =   975
   End
   Begin VB.HScrollBar hsrTransparency 
      Height          =   135
      LargeChange     =   50
      Left            =   6360
      Max             =   255
      TabIndex        =   45
      Top             =   5040
      Width           =   5535
   End
   Begin VB.Frame frameCode 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   43
      Top             =   6000
      Width           =   11775
      Begin VB.TextBox txtCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFF8E8&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   44
         Text            =   "frmMain.frx":0ECA
         Top             =   240
         Width           =   11535
      End
   End
   Begin VB.CommandButton cmdBuild 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFF8E8&
      Caption         =   "Build"
      Height          =   375
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   " Generate code "
      Top             =   5565
      Width           =   1695
   End
   Begin VB.CommandButton cmdBuildnCopy 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFF8E8&
      Caption         =   "Build n' Copy"
      Height          =   375
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   " Generate code and copy it to the clipboard "
      Top             =   5565
      Width           =   1815
   End
   Begin VB.TextBox txtVariable 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFF8E8&
      Height          =   285
      Left            =   6360
      TabIndex        =   40
      ToolTipText     =   " Enter the variable name to assign the MsgBox return value to "
      Top             =   4440
      Width           =   5535
   End
   Begin VB.CommandButton cmdBuildnCopynExit 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFF8E8&
      Caption         =   "Build n' Copy n' Exit"
      Height          =   375
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   " Copy code to clipboard & exit "
      Top             =   5565
      Width           =   1815
   End
   Begin VB.Frame frameType 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Type"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      Left            =   120
      TabIndex        =   19
      Top             =   4200
      Width           =   6135
      Begin VB.CheckBox chkModalType 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Neither"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   3240
         TabIndex        =   26
         ToolTipText     =   "Tooltip"
         Top             =   360
         Width           =   975
      End
      Begin VB.CheckBox chkReadR2L 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Make right to left readable (like Hebrew or Arabic)"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2040
         TabIndex        =   25
         Tag             =   "vbMsgBoxRtlReading"
         ToolTipText     =   "Tooltip"
         Top             =   1440
         Width           =   3975
      End
      Begin VB.CheckBox chkAlignRight 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Align text to the right"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Tag             =   "vbMsgBoxRight"
         ToolTipText     =   "Tooltip"
         Top             =   1440
         Width           =   1935
      End
      Begin VB.CheckBox chkForeground 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Set as the foreground window"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Tag             =   "VbMsgBoxSetForeground"
         ToolTipText     =   "Tooltip"
         Top             =   1080
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CheckBox chkHelp 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "'Help' button"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Tag             =   "vbMsgBoxHelpButton"
         ToolTipText     =   "Tooltip"
         Top             =   720
         Width           =   1335
      End
      Begin VB.CheckBox chkModalType 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "System modal"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   1800
         TabIndex        =   21
         Tag             =   "vbSystemModal"
         ToolTipText     =   "Tooltip"
         Top             =   360
         Width           =   1455
      End
      Begin VB.CheckBox chkModalType 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Application modal"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Tag             =   "vbApplicationModal"
         ToolTipText     =   "Tooltip"
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame frameButtonSet 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Button set"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   5160
      TabIndex        =   5
      Top             =   120
      Width           =   6735
      Begin VB.Frame frameDefaultButton 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Deafault Button"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   34
         Top             =   3120
         Width           =   6495
         Begin VB.CheckBox chkDefaultButton 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "Default button 4"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   3
            Left            =   4800
            TabIndex        =   38
            Tag             =   "vbDefaultButton4"
            ToolTipText     =   "Tooltip"
            Top             =   360
            Width           =   1575
         End
         Begin VB.CheckBox chkDefaultButton 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "Default button 3"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   3240
            TabIndex        =   37
            Tag             =   "vbDefaultButton3"
            ToolTipText     =   "Tooltip"
            Top             =   360
            Width           =   1575
         End
         Begin VB.CheckBox chkDefaultButton 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "Default button 2"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   1680
            TabIndex        =   36
            Tag             =   "vbDefaultButton2"
            ToolTipText     =   "Tooltip"
            Top             =   360
            Width           =   1575
         End
         Begin VB.CheckBox chkDefaultButton 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "Default button 1"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   35
            Tag             =   "vbDefaultButton1"
            ToolTipText     =   "Tooltip"
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.CheckBox chkButtonSet 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "vbOkOnly"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   33
         Tag             =   "vbOkOnly"
         ToolTipText     =   "Tooltip"
         Top             =   360
         Width           =   1215
      End
      Begin VB.CheckBox chkButtonSet 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "vbOkCancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   32
         Tag             =   "vbOkCancel"
         ToolTipText     =   "Tooltip"
         Top             =   825
         Width           =   1455
      End
      Begin VB.CheckBox chkButtonSet 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "vbAbortRetryIgnore"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   31
         Tag             =   "vbAbortRetryIgnore"
         ToolTipText     =   "Tooltip"
         Top             =   1290
         Width           =   2175
      End
      Begin VB.CheckBox chkButtonSet 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "vbYesNoCancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   30
         Tag             =   "vbYesNoCancel"
         ToolTipText     =   "Tooltip"
         Top             =   1755
         Width           =   1695
      End
      Begin VB.CheckBox chkButtonSet 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "vbYesNo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   29
         Tag             =   "vbYesNo"
         ToolTipText     =   "Tooltip"
         Top             =   2220
         Width           =   1095
      End
      Begin VB.CheckBox chkButtonSet 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "vbRetryCancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   28
         Tag             =   "vbRetryCancel"
         ToolTipText     =   "Tooltip"
         Top             =   2685
         Width           =   1695
      End
      Begin VB.CommandButton cmdButtonSet6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFF8E8&
         Caption         =   "Retry"
         Height          =   375
         Index           =   0
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Tooltip"
         Top             =   2595
         Width           =   1095
      End
      Begin VB.CommandButton cmdButtonSet6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFF8E8&
         Caption         =   "Cancel"
         Height          =   375
         Index           =   1
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Tooltip"
         Top             =   2595
         Width           =   1095
      End
      Begin VB.CommandButton cmdButtonSet5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFF8E8&
         Caption         =   "Yes"
         Height          =   375
         Index           =   0
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Tooltip"
         Top             =   2130
         Width           =   1095
      End
      Begin VB.CommandButton cmdButtonSet5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFF8E8&
         Caption         =   "No"
         Height          =   375
         Index           =   1
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Tooltip"
         Top             =   2130
         Width           =   1095
      End
      Begin VB.CommandButton cmdButtonSet4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFF8E8&
         Caption         =   "Yes"
         Height          =   375
         Index           =   0
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Tooltip"
         Top             =   1665
         Width           =   1095
      End
      Begin VB.CommandButton cmdButtonSet4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFF8E8&
         Caption         =   "No"
         Height          =   375
         Index           =   1
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Tooltip"
         Top             =   1665
         Width           =   1095
      End
      Begin VB.CommandButton cmdButtonSet4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFF8E8&
         Caption         =   "Cancel"
         Height          =   375
         Index           =   2
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Tooltip"
         Top             =   1665
         Width           =   1215
      End
      Begin VB.CommandButton cmdButtonSet3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFF8E8&
         Caption         =   "Abort"
         Height          =   375
         Index           =   0
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Tooltip"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdButtonSet3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFF8E8&
         Caption         =   "Retry"
         Height          =   375
         Index           =   1
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Tooltip"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdButtonSet3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFF8E8&
         Caption         =   "Ignore"
         Height          =   375
         Index           =   2
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Tooltip"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdButtonSet2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFF8E8&
         Caption         =   "Ok"
         Height          =   375
         Index           =   0
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Tooltip"
         Top             =   735
         Width           =   1095
      End
      Begin VB.CommandButton cmdButtonSet2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFF8E8&
         Caption         =   "Cancel"
         Height          =   375
         Index           =   1
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Tooltip"
         Top             =   735
         Width           =   1095
      End
      Begin VB.CommandButton cmdButtonSet1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFF8E8&
         Caption         =   "Ok"
         Height          =   375
         Index           =   0
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Tooltip"
         Top             =   270
         Width           =   1095
      End
   End
   Begin VB.Frame frameMessage 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Message"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   4935
      Begin VB.TextBox txtMessage 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFF8E8&
         Height          =   1695
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Text            =   "frmMain.frx":1110
         ToolTipText     =   " Message "
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Frame frameTitle 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4935
      Begin VB.TextBox txtTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFF8E8&
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Text            =   "Wife's servant!"
         ToolTipText     =   " Message box window title "
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Frame frameImage 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Icon"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   4935
      Begin VB.Shape shpIcon 
         BorderStyle     =   3  'Dot
         Height          =   780
         Left            =   90
         Top             =   1170
         Width           =   780
      End
      Begin VB.Image imgIcon 
         Height          =   720
         Index           =   0
         Left            =   120
         ToolTipText     =   " No icon "
         Top             =   240
         Width           =   720
      End
      Begin VB.Image imgIcon 
         Height          =   720
         Index           =   4
         Left            =   4080
         Picture         =   "frmMain.frx":12B5
         Tag             =   "vbExclamation"
         ToolTipText     =   " Exclamation icon "
         Top             =   240
         Width           =   720
      End
      Begin VB.Image imgIcon 
         Height          =   720
         Index           =   3
         Left            =   3090
         Picture         =   "frmMain.frx":217F
         Tag             =   "vbCritical"
         ToolTipText     =   " Critical error icon "
         Top             =   240
         Width           =   720
      End
      Begin VB.Image imgIcon 
         Height          =   720
         Index           =   2
         Left            =   2100
         Picture         =   "frmMain.frx":3049
         Tag             =   "vbQuestion"
         ToolTipText     =   " Question icon "
         Top             =   240
         Width           =   720
      End
      Begin VB.Image imgIcon 
         Height          =   720
         Index           =   1
         Left            =   1110
         Picture         =   "frmMain.frx":3F13
         Tag             =   "vbInformation"
         ToolTipText     =   " Information icon "
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Label lblRate 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Rate me..."
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   11115
      MouseIcon       =   "frmMain.frx":4DDD
      MousePointer    =   99  'Custom
      TabIndex        =   47
      Top             =   5280
      Width           =   780
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transperancy"
      Height          =   195
      Index           =   1
      Left            =   6360
      TabIndex        =   46
      Top             =   4800
      Width           =   990
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Set the return value to the variable below:"
      Height          =   195
      Index           =   0
      Left            =   6360
      TabIndex        =   39
      Top             =   4200
      Width           =   3075
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_NORMAL = 1
Const SW_SHOW = 1

'=========================================================================================================
'The following APIs are for form's transperancy effect, works under Win2K & XP only
'Please remove the API declarations and the constants and the EnableTransparancy subroutine
'to work under Win9x, also remove the hsrTransparency_Scroll sub
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const WS_EX_LAYERED = &H80000
'=========================================================================================================

Dim mIcon As Integer, mButtonSet As Integer, mDefaultButton As Integer, mModalType As Integer, mStr As String

Public Sub EnableTransparancy(ByVal hwnd As Long, Perc As Integer)
'If this routine causes error, simply bypass it, it has nothing to
'do with the original purpose of the utility
On Error Resume Next
SetWindowLong hwnd, GWL_EXSTYLE, GetWindowLong(hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
SetLayeredWindowAttributes hwnd, 0, Perc, LWA_ALPHA
End Sub

Private Sub cmdPreview_Click()
Dim mm
Dim tt

    If imgIcon(mIcon).Index = 1 Then mm = vbInformation
    If imgIcon(mIcon).Index = 4 Then mm = vbExclamation
    If imgIcon(mIcon).Index = 3 Then mm = vbCritical
    If imgIcon(mIcon).Index = 2 Then mm = vbQuestion

    'Sets the Keys
    If chkButtonSet(mButtonSet).Index = 0 Then tt = vbOKOnly
    If chkButtonSet(mButtonSet).Index = 1 Then tt = vbOKCancel
    If chkButtonSet(mButtonSet).Index = 2 Then tt = vbAbortRetryIgnore
    If chkButtonSet(mButtonSet).Index = 3 Then tt = vbYesNoCancel
    If chkButtonSet(mButtonSet).Index = 4 Then tt = vbYesNo
    If chkButtonSet(mButtonSet).Index = 5 Then tt = vbRetryCancel

MsgBox txtMessage, mm + tt, txtTitle
End Sub

Private Sub hsrTransparency_Change()
hsrTransparency_Scroll
End Sub

Private Sub hsrTransparency_Scroll()
'If this routine causes error, simply bypass it, it has nothing to
'do with the original purpose of the utility
On Error Resume Next
EnableTransparancy Me.hwnd, 255 - hsrTransparency.Value
End Sub

Private Sub chkButtonSet_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub

Dim a As Long
For a = 0 To 5
    If a <> Index Then chkButtonSet(a).Value = vbUnchecked Else chkButtonSet(a) = vbChecked
Next
mButtonSet = Index
End Sub

Private Sub chkDefaultButton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub

Dim a As Long
For a = 0 To 3
    If a <> Index Then chkDefaultButton(a).Value = vbUnchecked Else chkDefaultButton(a).Value = vbChecked
Next
mDefaultButton = Index
End Sub

Private Sub chkModalType_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub

Dim a As Long
For a = 0 To 2
    If a <> Index Then chkModalType(a).Value = vbUnchecked Else chkModalType(a).Value = vbChecked
Next
mModalType = Index
End Sub

Private Sub cmdBuild_Click()
If Trim(txtVariable) <> "" Then mStr = txtVariable & " = " Else mStr = ""
mStr = mStr & "MsgBox"
If Trim(txtVariable) <> "" Then mStr = mStr & "(" Else mStr = mStr & " "
mStr = mStr & ToVBString(txtMessage.Text) & ", "
If mIcon > 0 Then mStr = mStr & imgIcon(mIcon).Tag & " + "
mStr = mStr & chkButtonSet(mButtonSet).Tag
If mDefaultButton > 0 Then mStr = mStr & " + " & chkDefaultButton(mDefaultButton).Tag
If mModalType <> 2 Then mStr = mStr & " + " & chkModalType(mModalType).Tag
If chkHelp.Value = vbChecked Then mStr = mStr & " + " & chkHelp.Tag
If chkForeground.Value = vbChecked Then mStr = mStr & " + " & chkForeground.Tag
If chkAlignRight.Value = vbChecked Then mStr = mStr & " + " & chkAlignRight.Tag
If chkReadR2L.Value = vbChecked Then mStr = mStr & " + " & chkReadR2L.Tag
mStr = mStr & ", " & Chr(34) & txtTitle & Chr(34)
If Trim(txtVariable) <> "" Then mStr = mStr & ")"

txtCode = mStr
End Sub

Private Sub cmdBuildnCopy_Click()
cmdBuild_Click

Clipboard.Clear
Clipboard.SetText mStr
End Sub

Private Sub cmdBuildnCopynExit_Click()
cmdBuildnCopy_Click
Unload Me
End Sub

Private Sub cmdButtonSet1_Click(Index As Integer)
chkDefaultButton_MouseUp Index, 1, 0, 0, 0

chkButtonSet_MouseUp 0, 1, 0, 0, 0
End Sub

Private Sub cmdButtonSet2_Click(Index As Integer)
chkDefaultButton_MouseUp Index, 1, 0, 0, 0

chkButtonSet_MouseUp 1, 1, 0, 0, 0
End Sub

Private Sub cmdButtonSet3_Click(Index As Integer)
chkDefaultButton_MouseUp Index, 1, 0, 0, 0

chkButtonSet_MouseUp 2, 1, 0, 0, 0
End Sub

Private Sub cmdButtonSet4_Click(Index As Integer)
chkDefaultButton_MouseUp Index, 1, 0, 0, 0

chkButtonSet_MouseUp 3, 1, 0, 0, 0
End Sub

Private Sub cmdButtonSet5_Click(Index As Integer)
chkDefaultButton_MouseUp Index, 1, 0, 0, 0

chkButtonSet_MouseUp 4, 1, 0, 0, 0
End Sub

Private Sub cmdButtonSet6_Click(Index As Integer)
chkDefaultButton_MouseUp Index, 1, 0, 0, 0

chkButtonSet_MouseUp 5, 1, 0, 0, 0
End Sub

Private Sub Form_Load()
imgIcon_Click 1
chkButtonSet_MouseUp 0, 1, 0, 0, 0
chkDefaultButton_MouseUp 0, 1, 0, 0, 0
chkModalType_MouseUp 0, 1, 0, 0, 0

End Sub

Private Sub imgIcon_Click(Index As Integer)
shpIcon.Move imgIcon(Index).Left - 2 * Screen.TwipsPerPixelX, imgIcon(Index).Top - 2 * Screen.TwipsPerPixelY, imgIcon(Index).Width + 4 * Screen.TwipsPerPixelX, imgIcon(Index).Height + 4 * Screen.TwipsPerPixelY
mIcon = Index
End Sub

Private Sub lblRate_Click()
ShellExecute &O0, "Open", "http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=43585&lngWId=1", vbNullString, vbNullString, SW_NORMAL
End Sub

Private Sub lblRate_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then lblRate.Move lblRate.Left + 1, lblRate.Top + 1
End Sub

Private Sub lblRate_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then lblRate.Move lblRate.Left - 1, lblRate.Top - 1
End Sub

Private Function ToVBString(String2Conv As String) As String
'Convert special characters
String2Conv = Replace(String2Conv, Chr(34), Chr(34) & " & Chr(34) & " & Chr(34))
String2Conv = Replace(String2Conv, vbCrLf, Chr(34) & " & vbCrLf & " & Chr(34))
'Rip unwanted double quots and vbCrLf
String2Conv = Replace(String2Conv, Chr(34) & Chr(34), "")
String2Conv = Replace(String2Conv, "&  &", "&")

ToVBString = Chr(34) & String2Conv & Chr(34)
End Function

