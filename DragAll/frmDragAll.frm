VERSION 5.00
Begin VB.Form frmDragAll 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Real time dragging objects with very simple code"
   ClientHeight    =   5835
   ClientLeft      =   1815
   ClientTop       =   1830
   ClientWidth     =   9315
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDragAll.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   9315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   420
      Left            =   8100
      TabIndex        =   9
      Top             =   180
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Caption         =   "Moving Frame"
      Height          =   2670
      Left            =   6750
      TabIndex        =   3
      Top             =   2160
      Width           =   2130
      Begin VB.OptionButton Option1 
         Caption         =   "Move Me too"
         Height          =   285
         Left            =   225
         TabIndex        =   6
         Top             =   2070
         Width           =   1770
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Move Me"
         Height          =   285
         Left            =   135
         TabIndex        =   5
         Top             =   1170
         Width           =   1770
      End
      Begin VB.Label LblFrame 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Move us here"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   90
         TabIndex        =   4
         Top             =   315
         Width           =   1950
      End
   End
   Begin VB.PictureBox ImRemPro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2760
      Left            =   180
      Picture         =   "frmDragAll.frx":0E42
      ScaleHeight     =   2760
      ScaleWidth      =   6090
      TabIndex        =   0
      Top             =   765
      Width           =   6090
      Begin VB.Shape Shape1 
         Height          =   240
         Left            =   990
         Top             =   3645
         Width           =   3075
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Move Me Around"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   465
         Index           =   0
         Left            =   135
         TabIndex        =   1
         Top             =   135
         Width           =   3120
      End
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   $"frmDragAll.frx":34786
      Height          =   1140
      Left            =   90
      TabIndex        =   8
      Top             =   4590
      Width           =   3840
   End
   Begin VB.Label LblFormMove 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Move Entire Form"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   180
      TabIndex        =   7
      Top             =   180
      Width           =   8970
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Move Me Around"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   465
      Index           =   1
      Left            =   2745
      TabIndex        =   2
      Top             =   3690
      Width           =   3120
   End
End
Attribute VB_Name = "frmDragAll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PX As Integer
Dim PY As Integer




Private Sub Command1_Click()
End
End Sub

'code for Moving the picture
Private Sub ImRemPro_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then PX = X: PY = Y: ImRemPro.ZOrder 0
End Sub
Private Sub ImRemPro_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    ImRemPro.Move ImRemPro.Left + X - PX, ImRemPro.Top
    ImRemPro.Move ImRemPro.Left, ImRemPro.Top + Y - PY
End If
End Sub

'code for Moving the labels
Private Sub Label_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then PX = X: PY = Y: Label(Index).ZOrder 0
End Sub
Private Sub Label_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    Label(Index).Move Label(Index).Left + X - PX, Label(Index).Top
    Label(Index).Move Label(Index).Left, Label(Index).Top + Y - PY
End If
End Sub

'code for Moving the Window by the label
Private Sub LblFormMove_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then PX = X: PY = Y
End Sub
Private Sub LblFormMove_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    Me.Move Me.Left + X - PX, Me.Top
    Me.Move Me.Left, Me.Top + Y - PY
End If
End Sub

'code for Moving the Frame by the embedded label
Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Frame1.ZOrder 0
End Sub
Private Sub LblFrame_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then PX = X: PY = Y: Frame1.ZOrder 0
End Sub
Private Sub LblFrame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    Frame1.Move Frame1.Left + X - PX, Frame1.Top
    Frame1.Move Frame1.Left, Frame1.Top + Y - PY
End If
End Sub

'code for Moving the OptionBox
Private Sub Option1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then PX = X: PY = Y: Option1.ZOrder 0
End Sub

Private Sub Option1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    Option1.Move Option1.Left + X - PX, Option1.Top
    Option1.Move Option1.Left, Option1.Top + Y - PY
End If
End Sub

'code for Moving the Checkbox
Private Sub Check1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then PX = X: PY = Y: Check1.ZOrder 0
End Sub

Private Sub Check1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    Check1.Move Check1.Left + X - PX, Check1.Top
    Check1.Move Check1.Left, Check1.Top + Y - PY
End If
End Sub

