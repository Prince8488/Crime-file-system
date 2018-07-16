VERSION 5.00
Begin VB.Form frmChargeSheet 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Charge Sheet"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   8310
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H8000000E&
      Height          =   5070
      Left            =   1440
      TabIndex        =   46
      Top             =   960
      Width           =   6900
      Begin VB.CommandButton cmdBack 
         Caption         =   "Back"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3285
         TabIndex        =   23
         Top             =   3825
         Width           =   1005
      End
      Begin VB.CommandButton cmdFinish 
         Caption         =   "Finish"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4590
         TabIndex        =   22
         Top             =   3825
         Width           =   1005
      End
      Begin VB.TextBox txtWitnName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3120
         TabIndex        =   19
         Top             =   1425
         Width           =   2490
      End
      Begin VB.TextBox txtWitnAdd 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3120
         TabIndex        =   20
         Top             =   2182
         Width           =   2490
      End
      Begin VB.TextBox txtWitnOcc 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3120
         TabIndex        =   21
         Top             =   2940
         Width           =   2490
      End
      Begin VB.Label Label20 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Witness Details"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2280
         TabIndex        =   50
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label21 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1155
         TabIndex        =   49
         Top             =   1515
         Width           =   1290
      End
      Begin VB.Label Label22 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1155
         TabIndex        =   48
         Top             =   2190
         Width           =   1290
      End
      Begin VB.Label Label23 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Occupation"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1155
         TabIndex        =   47
         Top             =   2865
         Width           =   1290
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H8000000E&
      Height          =   6120
      Left            =   1080
      TabIndex        =   37
      Top             =   840
      Width           =   6900
      Begin VB.CommandButton cmdBack3 
         Caption         =   "Back"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2925
         TabIndex        =   53
         Top             =   5580
         Width           =   1365
      End
      Begin VB.ComboBox cmbAccuStat 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmChargeSheet.frx":0000
         Left            =   3870
         List            =   "frmChargeSheet.frx":000A
         TabIndex        =   16
         Text            =   "Custody"
         Top             =   4275
         Width           =   1965
      End
      Begin VB.CommandButton cmdNext3 
         Caption         =   "Next"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4455
         TabIndex        =   18
         Top             =   5580
         Width           =   1365
      End
      Begin VB.ComboBox cmbAccuSex 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmChargeSheet.frx":0021
         Left            =   3885
         List            =   "frmChargeSheet.frx":002B
         TabIndex        =   14
         Text            =   "Male"
         Top             =   2916
         Width           =   1965
      End
      Begin VB.TextBox txtAccuName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3885
         TabIndex        =   11
         Top             =   825
         Width           =   1965
      End
      Begin VB.TextBox txtAccuAdd 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3885
         TabIndex        =   12
         Top             =   1522
         Width           =   1965
      End
      Begin VB.TextBox txtAccuOcc 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3840
         TabIndex        =   13
         Top             =   2160
         Width           =   1965
      End
      Begin VB.TextBox txtAccuAge 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3885
         MaxLength       =   2
         TabIndex        =   15
         Top             =   3538
         Width           =   1965
      End
      Begin VB.TextBox txtAccuAct 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3885
         TabIndex        =   17
         Top             =   4935
         Width           =   1965
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Accused Details"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2145
         TabIndex        =   45
         Top             =   150
         Width           =   2415
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1380
         TabIndex        =   44
         Top             =   915
         Width           =   1440
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1320
         TabIndex        =   43
         Top             =   1560
         Width           =   1440
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Occupation"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1380
         TabIndex        =   42
         Top             =   2299
         Width           =   1440
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Sex"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1380
         TabIndex        =   41
         Top             =   2991
         Width           =   1440
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Age"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1380
         TabIndex        =   40
         Top             =   3683
         Width           =   1440
      End
      Begin VB.Label Label18 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Current Status"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1380
         TabIndex        =   39
         Top             =   4375
         Width           =   1440
      End
      Begin VB.Label Label19 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Action"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1380
         TabIndex        =   38
         Top             =   5070
         Width           =   1440
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000018&
      Height          =   4800
      Left            =   960
      TabIndex        =   31
      Top             =   840
      Width           =   6690
      Begin VB.CommandButton cmdBack2 
         Caption         =   "Back"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3240
         TabIndex        =   52
         Top             =   4155
         Width           =   960
      End
      Begin VB.CommandButton cmdNext2 
         Caption         =   "Next"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4365
         TabIndex        =   10
         Top             =   4140
         Width           =   960
      End
      Begin VB.TextBox txtInfoName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3375
         TabIndex        =   6
         Top             =   1275
         Width           =   1965
      End
      Begin VB.TextBox txtInfoAdd 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3375
         TabIndex        =   7
         Top             =   1905
         Width           =   1965
      End
      Begin VB.TextBox txtInfoOcc 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3375
         TabIndex        =   8
         Top             =   2535
         Width           =   1965
      End
      Begin VB.TextBox txtInfoPart 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3375
         TabIndex        =   9
         Top             =   3150
         Width           =   1965
      End
      Begin VB.Label Label24 
         BackColor       =   &H80000018&
         Caption         =   "(When/Where)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3600
         TabIndex        =   51
         Top             =   3555
         Width           =   1365
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000018&
         Caption         =   "Informat Status"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2355
         TabIndex        =   36
         Top             =   315
         Width           =   2265
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000018&
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   1080
         TabIndex        =   35
         Top             =   1200
         Width           =   1440
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000018&
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   1095
         TabIndex        =   34
         Top             =   1920
         Width           =   1440
      End
      Begin VB.Label Label10 
         BackColor       =   &H80000018&
         Caption         =   "Occupation"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   1095
         TabIndex        =   33
         Top             =   2550
         Width           =   1440
      End
      Begin VB.Label Label11 
         BackColor       =   &H80000018&
         Caption         =   "Particular"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   1095
         TabIndex        =   32
         Top             =   3240
         Width           =   1440
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Height          =   4800
      Left            =   840
      TabIndex        =   0
      Top             =   960
      Width           =   6690
      Begin VB.TextBox txtDate 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3540
         TabIndex        =   2
         Top             =   1800
         Width           =   2415
      End
      Begin VB.ComboBox cmbFirNo 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3555
         TabIndex        =   3
         Top             =   2520
         Width           =   2355
      End
      Begin VB.CommandButton cmdNext1 
         Caption         =   "Next"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4680
         TabIndex        =   5
         Top             =   4095
         Width           =   1275
      End
      Begin VB.TextBox txtPstatName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3540
         TabIndex        =   1
         Top             =   1095
         Width           =   2415
      End
      Begin VB.TextBox txtChrgNo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3540
         MaxLength       =   3
         TabIndex        =   24
         Top             =   375
         Width           =   2415
      End
      Begin VB.TextBox txtDist 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3555
         TabIndex        =   4
         Top             =   3240
         Width           =   2415
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000E&
         Caption         =   "District"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   465
         TabIndex        =   29
         Top             =   3300
         Width           =   2190
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000E&
         Caption         =   "FIR No:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   465
         TabIndex        =   28
         Top             =   2589
         Width           =   2190
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000E&
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   465
         TabIndex        =   27
         Top             =   1881
         Width           =   2190
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "Name Of  Police Station"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   465
         TabIndex        =   26
         Top             =   1173
         Width           =   2190
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "Charge Sheet No:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   480
         TabIndex        =   25
         Top             =   240
         Width           =   2190
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Charge Sheet"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3307
      TabIndex        =   30
      Top             =   270
      Width           =   2190
   End
End
Attribute VB_Name = "frmChargeSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnt As Integer
Dim X As Control

Private Sub cmbAccuStat_GotFocus()
Dim a As Integer
a = val(txtAccuAge.Text)

If (a < 18) Or (a >= 100) Then
MsgBox "Invalid age"
txtAccuAge.Text = " "
txtAccuAge.SetFocus
End If
End Sub

Private Sub cmdBack_Click()
Frame4.Visible = False
Frame3.Visible = True
End Sub

Private Sub cmdBack2_Click()
Frame2.Visible = False
Frame1.Visible = True
End Sub

Private Sub cmdBack3_Click()
Frame3.Visible = False
Frame2.Visible = True
End Sub

Private Sub cmdFinish_Click()
If (txtWitnAdd.Text = "" Or txtWitnName.Text = "" Or txtWitnOcc.Text = "") Then
    MsgBox "Missing Fields!!, Please fill up all", vbInformation, "Crime File system"
Else
    Set rs = con.Execute("select * from ChargeSheet where FirNo=" + cmbFirNo.Text + "")
    If (Not rs.EOF) Then
        MsgBox "Dupplication is not allowed, Please try another fir number", vbCritical, "Crime File system"
        Frame4.Visible = False
        Frame1.Visible = True
        cmbFirNo.SetFocus
    Else
        con.Execute ("insert into ChargeSheet values(" + txtChrgNo.Text + ",'" + txtPstatName.Text + "','" + txtDate.Text + _
        "'," + cmbFirNo.Text + ",'" + txtDist.Text + "','" + txtInfoName.Text + _
        "','" + txtInfoAdd.Text + "','" + txtInfoOcc.Text + "','" + txtInfoPart.Text + _
        "','" + txtAccuName.Text + "','" + txtAccuAdd.Text + "','" + cmbAccuSex.Text + _
        "'," + txtAccuAge.Text + ",'" + txtAccuOcc.Text + "','" + cmbAccuStat.Text + _
        "','" + txtAccuAct.Text + "','" + txtWitnName.Text + "','" + txtWitnAdd.Text + _
        "','" + txtWitnOcc.Text + "')")
        MsgBox "Record Added Successfully!!", vbInformation, "Crime File system"
        Frame4.Visible = False
        Frame1.Visible = True
        txtChrgNo.Text = txtChrgNo.Text + 1
      '  For Each x In Me.Controls
       '     If (x = TextBox) Then
        '        x.Text = ""
         '   End If
        'Next
    End If
End If
End Sub

Private Sub cmdNext1_Click()
If (txtChrgNo.Text = "" Or txtPstatName.Text = "" Or txtDate.Text = "" Or cmbFirNo.Text = "" Or txtDist.Text = "") Then
    MsgBox "Missing Fields!!, Please fill up all", vbInformation, "Crime File system"
Else
    Frame1.Visible = False
    Frame2.Visible = True
End If
End Sub

Private Sub cmdNext2_Click()
If (txtInfoAdd.Text = "" Or txtInfoName.Text = "" Or txtInfoOcc.Text = "" Or txtInfoPart.Text = "") Then
    MsgBox "Missing Fields!!, Please fill up all", vbInformation, "Crime File system"
Else
    Frame2.Visible = False
    Frame3.Visible = True
End If
End Sub

Private Sub cmdNext3_Click()
If (txtAccuAct.Text = "" Or txtAccuAdd.Text = "" Or txtAccuAge.Text = "" Or txtAccuName.Text = "" Or txtAccuOcc.Text = "") Then
    MsgBox "Missing Fields!!, Please fill up all", vbInformation, "Crime File system"
Else
    Frame3.Visible = False
    Frame4.Visible = True
End If
End Sub

Private Sub Form_Load()
Frame1.Visible = True
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False
dbconnection
Set rs = con.Execute("select max(ChrgShtNo) from ChargeSheet")
If (Not rs.EOF) Then
    cnt = rs(0)
    If (cnt = 0) Then
        cnt = 1
    Else
        cnt = cnt + 1
    End If
    txtChrgNo.Text = cnt
End If
rs.Close
Set rs = con.Execute("select Firno from FIR")
While (Not rs.EOF)
    cmbFirNo.AddItem rs(0)
    rs.MoveNext
Wend
rs.Close
End Sub

Private Sub txtAccuAct_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
KeyAscii = 0
MsgBox ("Please enter only alphabets")
End If
End Sub

Private Sub txtAccuAdd_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
KeyAscii = 0
MsgBox ("Please enter only alphabets")
End If
End Sub

Private Sub txtAccuAge_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57 Or KeyAscii = 32) Then
KeyAscii = 0
MsgBox ("Please enter only Numbers")
End If
End Sub

Private Sub txtAccuName_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
KeyAscii = 0
MsgBox ("Please enter only alphabets")
End If
End Sub

Private Sub txtAccuOcc_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
KeyAscii = 0
MsgBox ("Please enter only alphabets")
End If
End Sub

Private Sub txtChrgNo_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57 Or KeyAscii = 32) Then
KeyAscii = 0
MsgBox ("Please enter only Numbers")
End If
End Sub

Private Sub txtDate_Change()
txtDate.Text = Date
End Sub

Private Sub txtDist_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
KeyAscii = 0
MsgBox ("Please enter only alphabets")
End If
End Sub

Private Sub txtInfoAdd_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
KeyAscii = 0
MsgBox ("Please enter only alphabets")
End If
End Sub

Private Sub txtInfoName_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
KeyAscii = 0
MsgBox ("Please enter only alphabets")
End If
End Sub

Private Sub txtInfoOcc_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
KeyAscii = 0
MsgBox ("Please enter only alphabets")
End If
End Sub

Private Sub txtInfoPart_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
KeyAscii = 0
MsgBox ("Please enter only alphabets")
End If
End Sub

Private Sub txtPstatName_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
KeyAscii = 0
MsgBox ("Please enter only alphabets")
End If
End Sub

Private Sub txtWitnAdd_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
KeyAscii = 0
MsgBox ("Please enter only alphabets")
End If
End Sub

Private Sub txtWitnName_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
KeyAscii = 0
MsgBox ("Please enter only alphabets")
End If
End Sub

Private Sub txtWitnOcc_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
KeyAscii = 0
MsgBox ("Please enter only alphabets")
End If
End Sub
