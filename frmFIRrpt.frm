VERSION 5.00
Begin VB.Form frmFIRrpt 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FIR Report"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   7275
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   2295
      Left            =   1110
      TabIndex        =   0
      Top             =   1440
      Width           =   5055
      Begin VB.ComboBox cmbMno 
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
         Left            =   2040
         TabIndex        =   2
         Top             =   720
         Width           =   2295
      End
      Begin VB.CommandButton cmdShRpt 
         Caption         =   "Show Report"
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
         Left            =   1800
         TabIndex        =   1
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H00808080&
         Caption         =   "Select FIR No:"
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
         Left            =   360
         TabIndex        =   3
         Top             =   720
         Width           =   1815
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000014&
      Caption         =   "FIR"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3270
      TabIndex        =   4
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmFIRrpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X As Integer
Private Sub cmdShRpt_Click()
X = cmbMno.Text
    If (DataEnvironment1.rsCommand4.State = 1) Then
        DataEnvironment1.rsCommand4.Close
    Else
        DataEnvironment1.Command4 (X)
        Load DataReportFIR
        DataReportFIR.Show
    End If
End Sub

Private Sub Form_Load()
dbconnection
Set rs = con.Execute("select Firno from FIR")
While (Not rs.EOF)
    cmbMno.AddItem rs(0)
    rs.MoveNext
Wend
End Sub

