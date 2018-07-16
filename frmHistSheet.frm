VERSION 5.00
Begin VB.Form frmHistSheet 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "History Sheet"
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   9480
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Height          =   7710
      Left            =   1110
      TabIndex        =   0
      Top             =   472
      Width           =   7260
      Begin VB.ComboBox cmbCrimeType 
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
         ItemData        =   "frmHistSheet.frx":0000
         Left            =   3360
         List            =   "frmHistSheet.frx":000D
         TabIndex        =   15
         Top             =   2760
         Width           =   2175
      End
      Begin VB.TextBox txtDesc 
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
         Height          =   1065
         Left            =   3330
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   5265
         Width           =   3600
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
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
         Left            =   2175
         TabIndex        =   6
         Top             =   7035
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
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
         Left            =   3795
         TabIndex        =   5
         Top             =   7035
         Width           =   1215
      End
      Begin VB.ComboBox cmbPrisNo 
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
         Left            =   3330
         TabIndex        =   4
         Top             =   1080
         Width           =   2205
      End
      Begin VB.TextBox txtCrimeNo 
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
         Height          =   345
         Left            =   3330
         MaxLength       =   3
         TabIndex        =   3
         Top             =   1890
         Width           =   2205
      End
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
         Height          =   345
         Left            =   3330
         TabIndex        =   2
         Top             =   3585
         Width           =   2205
      End
      Begin VB.TextBox txtPlace 
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
         Height          =   345
         Left            =   3330
         TabIndex        =   1
         Top             =   4425
         Width           =   2205
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "         History  Sheet"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   2025
         TabIndex        =   13
         Top             =   135
         Width           =   3015
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   " Prisoners No:"
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
         Left            =   630
         TabIndex        =   12
         Top             =   1185
         Width           =   1590
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "Crime No:"
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
         Left            =   630
         TabIndex        =   11
         Top             =   1920
         Width           =   1590
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000E&
         Caption         =   "Type Of Crime"
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
         Left            =   630
         TabIndex        =   10
         Top             =   2790
         Width           =   1590
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000E&
         Caption         =   "Date Of Occurance"
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
         Left            =   630
         TabIndex        =   9
         Top             =   3615
         Width           =   1590
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000E&
         Caption         =   "Place Of Occurance"
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
         Left            =   630
         TabIndex        =   8
         Top             =   4485
         Width           =   2265
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000E&
         Caption         =   "Brief Description Of Case"
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
         Left            =   630
         TabIndex        =   7
         Top             =   5400
         Width           =   2400
      End
   End
End
Attribute VB_Name = "frmHistSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
If (txtCrimeNo.Text = "" Or cmbCrimeType.Text = "" Or txtDate.Text = "" Or txtDesc.Text = "" Or txtPlace.Text = "") Then
    MsgBox "Missing Fields", vbInformation, "CFS"
    ElseIf (Not IsNumeric(txtCrimeNo.Text)) Then
    MsgBox "Crime Number should be a number", vbCritical, "CFS"
Else
    Set rs = con.Execute("select * from History where CrimeNo=" + txtCrimeNo.Text + "")
    If (Not rs.EOF) Then
        MsgBox "Duplication of record is not allowded! Try another Crime number", vbCritical, "CFS"
        txtCrimeNo.Text = ""
        txtCrimeNo.SetFocus
    Else
    con.Execute ("insert into History values(" + cmbPrisNo.Text + "," + txtCrimeNo.Text + _
    ",'" + cmbCrimeType.Text + "','" + txtDate.Text + "','" + txtPlace.Text + "','" + txtDesc.Text + "')")
    con.Execute ("UPDATE PrisonersTemp set Status='Yes' where PrisonerNo=" + cmbPrisNo.Text + "")
    MsgBox "Recor Added successfully", vbInformation, "CFS"
    txtCrimeNo.Text = ""
    cmbCrimeType.Text = ""
    txtDate.Text = ""
    txtDesc.Text = ""
    txtPlace.Text = ""
    txtCrimeNo.SetFocus
    End If
End If
End Sub

Private Sub Form_Load()
dbconnection
Set rs = con.Execute("select PrisonerNo from PrisonersTemp where Status='No'")
While (Not rs.EOF)
    cmbPrisNo.AddItem rs(0)
    rs.MoveNext
Wend
rs.Close
End Sub

Private Sub txtCrimeNo_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57 Or KeyAscii = 32) Then
KeyAscii = 0
MsgBox ("Please enter only Numbers")
End If
End Sub

Private Sub txtDate_Change()
txtDate.Text = Date
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57 Or KeyAscii = 32) Then
KeyAscii = 0
MsgBox ("Please enter only Numbers")
End If
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
KeyAscii = 0
MsgBox ("Please enter only alphabets")
End If
End Sub

Private Sub txtPlace_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
KeyAscii = 0
MsgBox ("Please enter only alphabets")
End If
End Sub
