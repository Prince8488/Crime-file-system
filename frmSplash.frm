VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H80000009&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5085
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   5670
      TabIndex        =   7
      Top             =   4230
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   6975
      Top             =   4275
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000009&
      Height          =   330
      Left            =   225
      ScaleHeight     =   270
      ScaleWidth      =   6885
      TabIndex        =   6
      Top             =   4635
      Width           =   6945
      Begin VB.Image Image1 
         Height          =   285
         Left            =   0
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   0
         Width           =   465
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Height          =   4050
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   7080
      Begin VB.PictureBox Picture1 
         Height          =   2055
         Left            =   600
         Picture         =   "frmSplash.frx":2B60
         ScaleHeight     =   1995
         ScaleWidth      =   1875
         TabIndex        =   10
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label lblCopyright 
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright : Reserved"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5280
         TabIndex        =   2
         Top             =   3060
         Width           =   1575
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   " Warning: Copyright Protected"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   1
         Top             =   3660
         Width           =   6855
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version : 6.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   5430
         TabIndex        =   3
         Top             =   2700
         Width           =   1425
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Platform: Visual Basic"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   360
         Left            =   3555
         TabIndex        =   4
         Top             =   2340
         Width           =   3300
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Crime File System"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   765
         Left            =   1080
         TabIndex        =   5
         Top             =   540
         Width           =   5595
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
      Height          =   240
      Left            =   2025
      TabIndex        =   9
      Top             =   4320
      Width           =   3120
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "Loading Files....."
      Height          =   240
      Left            =   315
      TabIndex        =   8
      Top             =   4320
      Width           =   1590
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim x As Integer
Option Explicit

Private Sub Form_Load()
File1.FileName = App.Path
x = File1.ListCount
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Load frmMDI
frmMDI.Show

End Sub

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Unload Me
End Sub

Private Sub Timer1_Timer()
If (Image1.Left <= 6480) Then
    Image1.Left = Image1.Left + 100
Else
    Image1.Left = 0
End If
If (i <= x) Then
Label2.Caption = File1.List(i)
i = i + 1
Else
Load frmMDI
frmMDI.Show
Unload Me
End If
End Sub
