VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Collision Detection"
   ClientHeight    =   2550
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   3165
   LinkTopic       =   "Form1"
   ScaleHeight     =   2550
   ScaleWidth      =   3165
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      Caption         =   "Collision"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   615
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFC0&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   975
      Top             =   1028
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   240
      Top             =   360
      Width           =   1215
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***
'Very Simple Collision Detection
'Mostly Written by MDavis, with credit given to Rosh Mendis for the actual collision function
'***
'I'm fairly new to VB and I needed a collision detection for a small game I am writing.
'My web searches led me to complicated (to me, anyway:)) posts that had a lot more
'than just collision, which confused me.
'So, I wrote this to learn more for myself, and now that I understand it better,
'I thought I'd post it to possibly help someone else.
'***
'The intent here is for simple collision detection only.  Speed and precision
'were not really needed, nor considered.
Option Explicit

Public Function CollisionDetection(Object1 As Object, Object2 As Object) As Boolean
'Written by Rosh Mendis, via PSC post
'This function finds whether two objects have
'collided or not. Returns true or false
'
If Object1.Left < Object2.Left + Object2.Width And Object1.Left + Object1.Width > Object2.Left And Object1.Top < Object2.Top + Object2.Height And Object1.Top + Object1.Height > Object2.Top Then
    CollisionDetection = True
Else
    CollisionDetection = False
End If
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'Allows the user to use the keyboard cursor keys to move Shape1
Select Case KeyCode
    Case vbKeyLeft
        'Move Shape1 to the left
        Shape1.Left = Shape1.Left - 50
        If CollisionDetection(Shape1, Shape2) = True Then
            Label1.Visible = True
        Else
            Label1.Visible = False
        End If
    
    Case vbKeyRight
        'Move Shape1 to the right
        Shape1.Left = Shape1.Left + 50
        If CollisionDetection(Shape1, Shape2) = True Then
            Label1.Visible = True
        Else
            Label1.Visible = False
        End If
    
    Case vbKeyUp
        'Move Shape1 up
        Shape1.Top = Shape1.Top - 50
        If CollisionDetection(Shape1, Shape2) = True Then
            Label1.Visible = True
        Else
            Label1.Visible = False
        End If
    
    Case vbKeyDown
        'Move Shape1 down
        Shape1.Top = Shape1.Top + 50
        If CollisionDetection(Shape1, Shape2) = True Then
            Label1.Visible = True
        Else
            Label1.Visible = False
        End If
End Select
End Sub

Private Sub mnuHelp_Click()
'Message box for the Help menu item
Dim Message As String
Message = "Use the keyboard cursor keys to move the yellow box around the form.  "
Message = Message + "When the yellow box collides with the blue box, a Collision message appears."

MsgBox Message, vbOKOnly, "Help"
End Sub
