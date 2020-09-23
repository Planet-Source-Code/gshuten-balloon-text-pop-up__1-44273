VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   6660
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   615
      Left            =   4200
      TabIndex        =   4
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   5880
      TabIndex        =   3
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   960
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   1395
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   1
      Left            =   240
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   0
      Left            =   360
      Shape           =   3  'Circle
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "Text Here !"
      Height          =   195
      Left            =   720
      TabIndex        =   2
      Top             =   3360
      Width           =   795
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderStyle     =   6  'Inside Solid
      Height          =   495
      Index           =   0
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderStyle     =   6  'Inside Solid
      Height          =   495
      Index           =   1
      Left            =   720
      Shape           =   4  'Rounded Rectangle
      Top             =   3360
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Make_ballon Command1.Top, Command1.Left, Command1.Width, "Click if you can !", 1
End Sub

 
Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Make_ballon Command2.Top, Command2.Left, Command2.Width, "Groovy baby !", 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Hide_Ballon
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     Make_ballon Picture1.Top, Picture1.Left, Picture1.Width, "Hello World Again, Still way cool !", 0
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     Make_ballon Text1.Top, Text1.Left, Text1.Width, "Hello World, check this out ! way cool !", 0
End Sub


Private Function Make_ballon(ByVal lTop As Long, ByVal lLeft As Long, ByVal lWidth As Long, ByVal strCap As String, Optional lDirection As Long)

    'set text
    Label1.Caption = strCap
    Shape1(0).Width = Label1.Width + 80
    Shape1(1).Width = Label1.Width + 80
    
    
 If lDirection = 0 Then
    'set postion
    Shape1(0).Top = lTop
    Shape2(0).Top = lTop + 50
    Shape2(1).Top = lTop + 130
    
    Shape1(0).Left = lLeft + lWidth + Shape2(0).Width + Shape2(1).Width + 40
    Shape2(0).Left = lLeft + lWidth + 150
    Shape2(1).Left = lLeft + lWidth + 10
    
    Shape1(1).Top = Shape1(0).Top + 20
    Shape1(1).Left = Shape1(0).Left + 20
        
    Label1.Top = Shape1(0).Top + 25
    Label1.Left = Shape1(0).Left + 25 + 5
 
 Else

    Shape1(0).Top = lTop + lTop
    Shape2(0).Top = lTop - 50
    Shape2(1).Top = lTop - 130
    
    Shape1(0).Left = lLeft - lWidth - Shape2(0).Width - Shape2(1).Width - 100
    Shape2(0).Left = lLeft - lWidth - 150
    Shape2(1).Left = lLeft - lWidth - 10
    
    Shape1(1).Top = Shape1(0).Top + 20
    Shape1(1).Left = Shape1(0).Left + 20
        
    Label1.Top = Shape1(0).Top + 25
    Label1.Left = Shape1(0).Left + 25 + 5
 
 End If
 
 
 
   'make me seen !
    Shape1(0).Visible = True
    Shape1(0).Visible = True
    Shape1(1).Visible = True
    Shape1(1).Visible = True
    Shape2(0).Visible = True
    Shape2(1).Visible = True
    Label1.Visible = True
    Label1.Visible = True

End Function

Private Function Hide_Ballon()
    Shape1(0).Visible = False
    Shape1(0).Visible = False
    Shape1(1).Visible = False
    Shape1(1).Visible = False
    Shape2(0).Visible = False
    Shape2(1).Visible = False
    Label1.Visible = False
    Label1.Visible = False
End Function
