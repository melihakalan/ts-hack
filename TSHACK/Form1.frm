VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "KoJD TS Hack 1745 v1.5"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2775
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   162
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   2775
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frHack 
      Caption         =   "Unlimited - Free Transformation"
      Height          =   3900
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   2535
      Begin VB.CommandButton btnHuge 
         Caption         =   "Tiny"
         Height          =   285
         Left            =   1320
         TabIndex        =   11
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CommandButton btnTiny 
         Caption         =   "Huge"
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CheckBox chkTiny 
         Caption         =   "Huge"
         Height          =   195
         Left            =   840
         TabIndex        =   9
         Top             =   2400
         Width           =   735
      End
      Begin VB.CheckBox chkHuge 
         Caption         =   "Tiny"
         Height          =   195
         Left            =   1680
         TabIndex        =   8
         Top             =   2400
         Width           =   735
      End
      Begin VB.CommandButton btnDel 
         Caption         =   "Delete ALL (Normal)"
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   3480
         Width           =   2295
      End
      Begin VB.CommandButton btnTS 
         Caption         =   "Transform"
         Default         =   -1  'True
         Height          =   400
         Left            =   120
         TabIndex        =   5
         Top             =   3000
         Width           =   2295
      End
      Begin VB.ListBox lstTS 
         Height          =   2010
         ItemData        =   "Form1.frx":0000
         Left            =   120
         List            =   "Form1.frx":00BE
         TabIndex        =   4
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Auto:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   2400
         Width           =   450
      End
   End
   Begin VB.Frame frAttach 
      Caption         =   "Attach"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2535
      Begin VB.CommandButton btnGO 
         Caption         =   "GO"
         Height          =   285
         Left            =   1800
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtTitle 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Text            =   "Knight OnLine Client"
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Timer TimerOfTheKoJD 
      Interval        =   50
      Left            =   2400
      Top             =   4680
   End
   Begin VB.Label LabelOfTheKOJD2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "KoJD"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   150
      MousePointer    =   2  'Cross
      TabIndex        =   13
      Top             =   4680
      Width           =   2505
   End
   Begin VB.Label LabelOfTheKOJD 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Onlinehile.com - Snoxd.net"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   150
      MousePointer    =   2  'Cross
      TabIndex        =   6
      Top             =   4890
      Width           =   2505
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Coloring
Public Coloring2

Private Sub btnDel_Click()
SendAPacket ("290301")
WriteLong KO_ADR_CHR + KO_OFF_SWIFT, 16256
End Sub

Private Sub btnGO_Click()
KO_TITLE = txtTitle.Text
LoadOffsets
If AttachKO = False Then
Exit Sub
End If
KO_ADR_CHR = ReadLong(KO_PTR_CHR)
KO_ADR_DLG = ReadLong(KO_PTR_DLG)
frHack.Enabled = True
frAttach.Enabled = False
End Sub

Private Sub btnHuge_Click()
SendAPacket ("290303")
End Sub

Private Sub btnTiny_Click()
SendAPacket ("290302")
End Sub

Private Sub btnTS_Click()
WriteLong KO_ADR_CHR + KO_OFF_SWIFT, 16256

Select Case lstTS.ListIndex

Case 0
SendAPacket "2903427E07"
Case 1
SendAPacket "2903397E07"
Case 2
SendAPacket "2903023507"
Case 3
SendAPacket "2903417E07"
Case 4
SendAPacket "2903407E07"
Case 5
SendAPacket "29033e7E07"
Case 6
SendAPacket "29033f7E07"
Case 7
SendAPacket "2903103407"
Case 8
SendAPacket "2903b83407"
Case 9
SendAPacket "2903a63407"
Case 10
SendAPacket "2903b03407"
Case 11
SendAPacket "2903c43407"
Case 12
SendAPacket "2903d03407"
Case 13
SendAPacket "2903d23407"
Case 14
SendAPacket "2903d43407"
Case 15
SendAPacket "2903d83407"
Case 16
SendAPacket "2903e23407"
Case 17
SendAPacket "2903ec3407"
Case 18
SendAPacket "2903f63407"
Case 19
SendAPacket "2903063407"
Case 20
SendAPacket "29032e3407"
Case 21
SendAPacket "2903383407"
Case 22
SendAPacket "29034c3407"
Case 23
SendAPacket "2903443407"
Case 24
SendAPacket "2903603407"
Case 25
SendAPacket "29036a3407"
Case 26
SendAPacket "2903803407"
Case 27
SendAPacket "29038a3407"
Case 28
SendAPacket "29039c3407"
Case 29
SendAPacket "2903943407"
Case 30
SendAPacket "2903e23407"
Case 31
SendAPacket "2903334b07"
Case 32
SendAPacket "2903344b07"
Case 33
SendAPacket "2903354b07"
Case 34
SendAPacket "2903364b07"
Case 35
SendAPacket "2903D43307"
Case 36
SendAPacket "2903ca3307"
Case 37
SendAPacket "2903de3307"
Case 38
SendAPacket "2903e83307"
Case 39
SendAPacket "2903f23307"
Case 40
SendAPacket "29033c3007"
Case 41
SendAPacket "2903663307"
Case 42
SendAPacket "2903703307"
Case 43
SendAPacket "2903023507"
SendAPacket "290302"
Case 45
SendAPacket "29031da307"
Case 46
SendAPacket "290319a307"
Case 47
SendAPacket "29031fa307"
Case 48
SendAPacket "29031ea307"
Case 49
SendAPacket "29031aa307"
Case 50
SendAPacket "29031ca307"
Case 51
SendAPacket "29031ba307"
Case 52
SendAPacket "290320a307"
Case 54
SendAPacket "2903d1dd06"
WriteLong KO_ADR_CHR + KO_OFF_SWIFT, 16100
Case 55
SendAPacket "2903d2dd06"
WriteLong KO_ADR_CHR + KO_OFF_SWIFT, 16100
Case 56
SendAPacket "2903d3dd06"
WriteLong KO_ADR_CHR + KO_OFF_SWIFT, 16100
Case 57
SendAPacket "2903d4dd06"
WriteLong KO_ADR_CHR + KO_OFF_SWIFT, 16100
Case 58
SendAPacket "2903e1dd06"
WriteLong KO_ADR_CHR + KO_OFF_SWIFT, 16100
Case 59
SendAPacket "2903e2dd06"
WriteLong KO_ADR_CHR + KO_OFF_SWIFT, 16100
Case 60
SendAPacket "2903e3dd06"
WriteLong KO_ADR_CHR + KO_OFF_SWIFT, 16100
Case 61
SendAPacket "2903e4dd06"
WriteLong KO_ADR_CHR + KO_OFF_SWIFT, 16100
Case Else

End Select

If chkTiny.Value = 1 Then SendAPacket ("290302")
If chkHuge.Value = 1 Then SendAPacket ("290303")

End Sub



Private Sub chkHuge_Click()
If chkHuge.Value = 1 Then chkTiny.Value = 0
End Sub

Private Sub chkTiny_Click()
If chkTiny.Value = 1 Then chkHuge.Value = 0
End Sub

Private Sub Form_Load()
Coloring = 0
Coloring2 = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
If webWebOnExit <> 0 Then
Shell "explorer " & webWebOnExit
End If
End
End Sub



Private Sub lstTS_Click()
btnTS.Default = True
End Sub

Private Sub TimerOfTheKoJD_Timer()
Coloring = Coloring + 1
Coloring2 = Coloring2 + 1

Select Case Coloring
Case 1
LabelOfTheKOJD.ForeColor = &HE0E0E0
LabelOfTheKOJD.BackColor = &H404040
Case 2
LabelOfTheKOJD.ForeColor = &HC0C0C0
LabelOfTheKOJD.BackColor = &H808080
Case 3
LabelOfTheKOJD.ForeColor = &H808080
LabelOfTheKOJD.BackColor = &HC0C0C0
Case 4
LabelOfTheKOJD.ForeColor = &H404040
LabelOfTheKOJD.BackColor = &HE0E0E0
Case 5
LabelOfTheKOJD.ForeColor = &H0
LabelOfTheKOJD.BackColor = &HFFFFFF
Case 6
LabelOfTheKOJD.ForeColor = &H404040
LabelOfTheKOJD.BackColor = &HE0E0E0
Case 7
LabelOfTheKOJD.ForeColor = &H808080
LabelOfTheKOJD.BackColor = &HC0C0C0
Case 8
LabelOfTheKOJD.ForeColor = &HC0C0C0
LabelOfTheKOJD.BackColor = &H808080
Case 9
LabelOfTheKOJD.ForeColor = &HE0E0E0
LabelOfTheKOJD.BackColor = &H404040
Case 10
LabelOfTheKOJD.ForeColor = &HFFFFFF
LabelOfTheKOJD.BackColor = &H0
Coloring = 0
End Select

Select Case Coloring2
Case 1
LabelOfTheKOJD2.ForeColor = &HE0E0E0
LabelOfTheKOJD2.BackColor = &H404040
Case 2
LabelOfTheKOJD2.ForeColor = &HC0C0C0
LabelOfTheKOJD2.BackColor = &H808080
Case 3
LabelOfTheKOJD2.ForeColor = &H808080
LabelOfTheKOJD2.BackColor = &HC0C0C0
Case 4
LabelOfTheKOJD2.ForeColor = &H404040
LabelOfTheKOJD2.BackColor = &HE0E0E0
Case 5
LabelOfTheKOJD2.ForeColor = &H0
LabelOfTheKOJD2.BackColor = &HFFFFFF
Case 6
LabelOfTheKOJD2.ForeColor = &H404040
LabelOfTheKOJD2.BackColor = &HE0E0E0
Case 7
LabelOfTheKOJD2.ForeColor = &H808080
LabelOfTheKOJD2.BackColor = &HC0C0C0
Case 8
LabelOfTheKOJD2.ForeColor = &HC0C0C0
LabelOfTheKOJD2.BackColor = &H808080
Case 9
LabelOfTheKOJD2.ForeColor = &HE0E0E0
LabelOfTheKOJD2.BackColor = &H404040
Case 10
LabelOfTheKOJD2.ForeColor = &HFFFFFF
LabelOfTheKOJD2.BackColor = &H0
Coloring2 = 0
End Select

End Sub
