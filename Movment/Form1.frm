VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   Caption         =   "Character To Mouse X, Y Movement"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9075
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   500
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   605
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5310
      TabIndex        =   0
      Text            =   "3"
      Top             =   60
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   120
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Movement Speed:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   240
      Left            =   3270
      TabIndex        =   1
      Top             =   120
      Width           =   1890
   End
   Begin VB.Shape Shape1 
      FillStyle       =   6  'Cross
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   9135
   End
   Begin VB.Image Image1 
      Height          =   285
      Left            =   4455
      Picture         =   "Form1.frx":08CA
      Top             =   3608
      Width           =   165
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////
'// Author: Derek Skeba               //
'// Email: frillyozz@comcast.net      //
'//                                   //
'// Description: Learn how to         //
'// move objects to any position      //
'// on the form, just by clicking     //
'// there.                            //
'//                                   //
'// *This code is a small portion     //
'// of code from my current project.  //
'// Which is a Realtime War Strategy, //
'// like 'Command And Conquer'.       //
'///////////////////////////////////////

'Basic declarations
Dim movex As Integer
Dim movey As Integer
Dim active As Boolean
Dim speedx As Integer
Dim speedy As Integer
Dim section As Single

Private Sub Form_Load()

'Set active to false, program knows that the character is not already moving
active = False

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If active = False Then

    If X > Image1.Left And Y > Image1.Top Then
        'Section 1 is would be when user clicks under and to the right of the character
        section = 1
    End If
    
    If X < Image1.Left And Y > Image1.Top Then
        'Section 2 is would be when user clicks under and to the left of the character
        section = 2
    End If
    
    If X < Image1.Left And Y < Image1.Top Then
        'Section 3 is would be when user clicks over and to the left of the character
        section = 3
    End If
    
    If X > Image1.Left And Y < Image1.Top Then
        'Section 4 is would be when user clicks over and to the right of the character
        section = 4
    End If
    
    'Set movex and movey to whatever mouse's X and Y coordinates are at the current time
    movex = X
    movey = Y
    
    'I put this here, just incase user puts nothing in the textbox, this way it will just skip this part and not move at all instead of bringing up an error
    On Error Resume Next
    speedy = Text1.Text
    speedx = Text1.Text
    
    'Set Timer1's intervel to 1
    Timer1.Interval = 1
    
End If

End Sub

Private Sub Timer1_Timer()

'Make the select case work with section variable
Select Case section
    
    Case 1
        'Statements to move the character, and stop it when it gets to the target position
        Image1.Left = Image1.Left + speedx
        Image1.Top = Image1.Top + speedy
        If Image1.Left >= movex Then
        speedx = 0
        End If
        If Image1.Top >= movey Then
        speedy = 0
        End If
        If speedy = 0 And speedx = 0 Then
        active = False
        Timer1.Interval = 0
        End If
        
    Case 2
        'Statements to move the character, and stop it when it gets to the target position
        Image1.Left = Image1.Left - speedx
        Image1.Top = Image1.Top + speedy
        If Image1.Left <= movex Then
        speedx = 0
        End If
        If Image1.Top >= movey Then
        speedy = 0
        End If
        If speedy = 0 And speedx = 0 Then
        active = False
        Timer1.Interval = 0
        End If
        
    Case 3
        'Statements to move the character, and stop it when it gets to the target position
        Image1.Left = Image1.Left - speedx
        Image1.Top = Image1.Top - speedy
        If Image1.Left <= movex Then
        speedx = 0
        End If
        If Image1.Top <= movey Then
        speedy = 0
        End If
        If speedy = 0 And speedx = 0 Then
        active = False
        Timer1.Interval = 0
        End If
        
    Case 4
        'Statements to move the character, and stop it when it gets to the target position
        Image1.Left = Image1.Left + speedx
        Image1.Top = Image1.Top - speedy
        If Image1.Left >= movex Then
        speedx = 0
        End If
        If Image1.Top <= movey Then
        speedy = 0
        End If
        If speedy = 0 And speedx = 0 Then
        active = False
        Timer1.Interval = 0
        End If
        
End Select

End Sub
