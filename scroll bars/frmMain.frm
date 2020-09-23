VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Scroll bars"
   ClientHeight    =   4080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6225
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   6225
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar VScroll1 
      Height          =   3135
      LargeChange     =   100
      Left            =   3720
      Max             =   2990
      Min             =   130
      SmallChange     =   10
      TabIndex        =   5
      Top             =   120
      Value           =   130
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   100
      Left            =   120
      Max             =   3230
      Min             =   30
      SmallChange     =   10
      TabIndex        =   4
      Top             =   3360
      Value           =   30
      Width           =   3495
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Top and Left"
      ForeColor       =   &H00000000&
      Height          =   1335
      Left            =   4560
      TabIndex        =   1
      Top             =   360
      Width           =   1215
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      Begin VB.Shape Shape1 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         Height          =   255
         Left            =   30
         Shape           =   4  'Rounded Rectangle
         Top             =   130
         Width           =   255
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   1
      Left            =   4485
      TabIndex        =   7
      Top             =   3315
      Width           =   1245
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   375
      Index           =   0
      Left            =   4440
      Shape           =   4  'Rounded Rectangle
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "About"
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   0
      Left            =   4530
      TabIndex        =   6
      Top             =   2835
      Width           =   1155
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   375
      Index           =   1
      Left            =   4440
      Shape           =   4  'Rounded Rectangle
      Top             =   2760
      Width           =   1335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3(0).FontBold = False
Label3(1).FontBold = False
End Sub

Private Sub HScroll1_Scroll()
    change
End Sub

Private Sub Label3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3(Index).ForeColor = vbBlack
End Sub

Private Sub Label3_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3(Index).FontBold = True
End Sub

Private Sub Label3_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3(Index).ForeColor = &H404040
Select Case Index
Case 0
MsgBox "Created by Petar Palasek & Igor Canadi"
Case 1
Unload Me
End Select
End Sub

Private Sub VScroll1_Scroll()
    change
End Sub
Private Sub HScroll1_change()
    change
End Sub

Private Sub VScroll1_change()
    change
End Sub

Private Sub change() 'this is important for scrolling the shape
    Shape1.Left = HScroll1.Value
    Shape1.Top = VScroll1.Value
    Label1.Caption = "Top: " & VScroll1.Value - VScroll1.Min
    Label2.Caption = "Left: " & HScroll1.Value - HScroll1.Min
End Sub
