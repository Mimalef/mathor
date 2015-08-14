VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   6540
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   5280
      Top             =   1920
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   40
      Text            =   "0"
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   38
      Text            =   "0"
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   35
      Text            =   "0"
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   33
      Text            =   "Player 1"
      Top             =   120
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   5760
      Top             =   1920
   End
   Begin VB.CommandButton Command17 
      Caption         =   "?"
      Height          =   855
      Left            =   3120
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton Command16 
      Caption         =   "?"
      Height          =   855
      Left            =   2160
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton Command15 
      Caption         =   "?"
      Height          =   855
      Left            =   1200
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton Command14 
      Caption         =   "?"
      Height          =   855
      Left            =   240
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton Command13 
      Caption         =   "?"
      Height          =   855
      Left            =   3120
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command12 
      Caption         =   "?"
      Height          =   855
      Left            =   2160
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command11 
      Caption         =   "?"
      Height          =   855
      Left            =   1200
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command10 
      Caption         =   "?"
      Height          =   855
      Left            =   240
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command9 
      Caption         =   "?"
      Height          =   855
      Left            =   3120
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "?"
      Height          =   855
      Left            =   2160
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "?"
      Height          =   855
      Left            =   1200
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      Caption         =   "?"
      Height          =   855
      Left            =   240
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "?"
      Height          =   855
      Left            =   3120
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "?"
      Height          =   855
      Left            =   2160
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "?"
      Height          =   855
      Left            =   1200
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "?"
      Height          =   855
      Left            =   240
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   120
      Width           =   855
   End
   Begin VB.PictureBox Picture16 
      Height          =   855
      Left            =   3120
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3000
      Width           =   855
   End
   Begin VB.PictureBox Picture15 
      Height          =   855
      Left            =   2160
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3000
      Width           =   855
   End
   Begin VB.PictureBox Picture14 
      Height          =   855
      Left            =   1200
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3000
      Width           =   855
   End
   Begin VB.PictureBox Picture13 
      Height          =   855
      Left            =   240
      Picture         =   "Form1.frx":22DA
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3000
      Width           =   855
   End
   Begin VB.PictureBox Picture12 
      Height          =   855
      Left            =   3120
      Picture         =   "Form1.frx":45B4
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2040
      Width           =   855
   End
   Begin VB.PictureBox Picture11 
      Height          =   855
      Left            =   2160
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2040
      Width           =   855
   End
   Begin VB.PictureBox Picture10 
      Height          =   855
      Left            =   1200
      Picture         =   "Form1.frx":688E
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2040
      Width           =   855
   End
   Begin VB.PictureBox Picture9 
      Height          =   855
      Left            =   240
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2040
      Width           =   855
   End
   Begin VB.PictureBox Picture8 
      Height          =   855
      Left            =   3120
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1080
      Width           =   855
   End
   Begin VB.PictureBox Picture7 
      Height          =   855
      Left            =   2160
      Picture         =   "Form1.frx":8B68
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1080
      Width           =   855
   End
   Begin VB.PictureBox Picture6 
      Height          =   855
      Left            =   1200
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1080
      Width           =   855
   End
   Begin VB.PictureBox Picture5 
      Height          =   855
      Left            =   240
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1080
      Width           =   855
   End
   Begin VB.PictureBox Picture4 
      Height          =   855
      Left            =   3120
      Picture         =   "Form1.frx":AE42
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   855
   End
   Begin VB.PictureBox Picture3 
      Height          =   855
      Left            =   2160
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   855
   End
   Begin VB.PictureBox Picture2 
      Height          =   855
      Left            =   1200
      Picture         =   "Form1.frx":D11C
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   240
      Picture         =   "Form1.frx":F3F6
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Player 2:"
      Height          =   255
      Left            =   4200
      TabIndex        =   39
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Player 1:"
      Height          =   255
      Left            =   4200
      TabIndex        =   37
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Scores:"
      Height          =   255
      Left            =   4200
      TabIndex        =   36
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Time:"
      Height          =   255
      Left            =   4200
      TabIndex        =   34
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Turn:"
      Height          =   255
      Left            =   4200
      TabIndex        =   32
      Top             =   120
      Width           =   735
   End
   Begin VB.Line Line1 
      X1              =   4080
      X2              =   4080
      Y1              =   120
      Y2              =   3960
   End
   Begin VB.Menu gg 
      Caption         =   "game"
      Begin VB.Menu ng 
         Caption         =   "new game"
      End
      Begin VB.Menu og 
         Caption         =   "open game"
      End
      Begin VB.Menu sg 
         Caption         =   "save game"
      End
      Begin VB.Menu ex 
         Caption         =   "exit"
      End
   End
   Begin VB.Menu hh 
      Caption         =   "help"
      Begin VB.Menu hw 
         Caption         =   "view help"
      End
      Begin VB.Menu ho 
         Caption         =   "about"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim click
Dim player
Dim pic1
Dim pic2
Dim btn2
Dim btn3
Dim btn4
Dim btn5
Dim btn6
Dim btn7
Dim btn8
Dim btn9
Dim btn10
Dim btn11
Dim btn12
Dim btn13
Dim btn14
Dim btn15
Dim btn16
Dim btn17
Dim btn01
Dim btn02

Sub check_click(btn)
    Call find_pic(btn.Name)
    click = click + 1
    
    If pic1 <> 0 And pic2 <> 0 Then
        If pic1 = pic2 Then
            If player = "Player 1" Then
                Text3 = Text3 + 1
            Else
                Text4 = Text4 + 1
            End If
            
            Call delete_button
        Else
            Call change_player
        End If
        
        Text1 = player
        Timer1.Enabled = True
        
        pic1 = 0
        pic2 = 0
        click = 0
    End If
End Sub

Sub find_pic(btn)
    If pic1 = 0 Then
        Select Case (btn)
            Case "Command2"
                pic1 = Picture1
                btn01 = 2
            Case "Command3"
                pic1 = Picture2
                btn01 = 3
            Case "Command4"
                pic1 = Picture3
                btn01 = 4
            Case "Command5"
                pic1 = Picture4
                btn01 = 5
            Case "Command6"
                pic1 = Picture5
                btn01 = 6
            Case "Command7"
                pic1 = Picture6
                btn01 = 7
            Case "Command8"
                pic1 = Picture7
                btn01 = 8
            Case "Command9"
                pic1 = Picture8
                btn01 = 9
            Case "Command10"
                pic1 = Picture9
                btn01 = 10
            Case "Command11"
                pic1 = Picture10
                btn01 = 11
            Case "Command12"
                pic1 = Picture11
                btn01 = 12
            Case "Command13"
                pic1 = Picture12
                btn01 = 13
            Case "Command14"
                pic1 = Picture13
                btn01 = 14
            Case "Command15"
                pic1 = Picture14
                btn01 = 15
            Case "Command16"
                pic1 = Picture15
                btn01 = 16
            Case "Command17"
                pic1 = Picture16
                btn01 = 17
        End Select
    Else
        Select Case (btn)
            Case "Command2"
                pic2 = Picture1
                btn02 = 2
            Case "Command3"
                pic2 = Picture2
                btn02 = 3
            Case "Command4"
                pic2 = Picture3
                btn02 = 4
            Case "Command5"
                pic2 = Picture4
                btn02 = 5
            Case "Command6"
                pic2 = Picture5
                btn02 = 6
            Case "Command7"
                pic2 = Picture6
                btn02 = 7
            Case "Command8"
                pic2 = Picture7
                btn02 = 8
            Case "Command9"
                pic2 = Picture8
                btn02 = 9
            Case "Command10"
                pic2 = Picture9
                btn02 = 10
            Case "Command11"
                pic2 = Picture10
                btn02 = 11
            Case "Command12"
                pic2 = Picture11
                btn02 = 12
            Case "Command13"
                pic2 = Picture12
                btn02 = 13
            Case "Command14"
                pic2 = Picture13
                btn02 = 14
            Case "Command15"
                pic2 = Picture14
                btn02 = 15
            Case "Command16"
                pic2 = Picture15
                btn02 = 16
            Case "Command17"
                pic2 = Picture16
                btn02 = 17
        End Select
    End If
End Sub

Sub change_player()
    If player = "Player 1" Then
        player = "Player 2"
    Else
        player = "Player 1"
    End If
End Sub

Sub delete_button()
    Select Case btn01
        Case 2
            btn2 = 1
        Case 3
            btn3 = 1
        Case 4
            btn4 = 1
        Case 5
            btn5 = 1
        Case 6
            btn6 = 1
        Case 7
            btn7 = 1
        Case 8
            btn8 = 1
        Case 9
            btn9 = 1
        Case 10
            btn10 = 1
        Case 11
            btn11 = 1
        Case 12
            btn12 = 1
        Case 13
            btn13 = 1
        Case 14
            btn14 = 1
        Case 15
            btn15 = 1
        Case 16
            btn16 = 1
        Case 17
            btn17 = 1
    End Select
    
    Select Case btn02
        Case 2
            btn2 = 1
        Case 3
            btn3 = 1
        Case 4
            btn4 = 1
        Case 5
            btn5 = 1
        Case 6
            btn6 = 1
        Case 7
            btn7 = 1
        Case 8
            btn8 = 1
        Case 9
            btn9 = 1
        Case 10
            btn10 = 1
        Case 11
            btn11 = 1
        Case 12
            btn12 = 1
        Case 13
            btn13 = 1
        Case 14
            btn14 = 1
        Case 15
            btn15 = 1
        Case 16
            btn16 = 1
        Case 17
            btn17 = 1
    End Select
End Sub

Private Sub Command1_Click()
    Text1 = "Player 1"
    Text2 = 0
    Text3 = 0
    Text4 = 0
    pic1 = 0
    pic2 = 0
    click = 0
    btn2 = 0
    btn3 = 0
    btn4 = 0
    btn5 = 0
    btn6 = 0
    btn7 = 0
    btn8 = 0
    btn9 = 0
    btn10 = 0
    btn11 = 0
    btn12 = 0
    btn13 = 0
    btn14 = 0
    btn15 = 0
    btn16 = 0
    btn17 = 0
    btn01 = 0
    btn02 = 0
    player = "Player 1"
    Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
    Command2.Visible = False
    Call check_click(Command2)
End Sub

Private Sub Command3_Click()
    Command3.Visible = False
    Call check_click(Command3)
End Sub

Private Sub Command4_Click()
    Command4.Visible = False
    Call check_click(Command4)
End Sub

Private Sub Command5_Click()
    Command5.Visible = False
    Call check_click(Command5)
End Sub

Private Sub Command6_Click()
    Command6.Visible = False
    Call check_click(Command6)
End Sub

Private Sub Command7_Click()
    Command7.Visible = False
    Call check_click(Command7)
End Sub

Private Sub Command8_Click()
    Command8.Visible = False
    Call check_click(Command8)
End Sub

Private Sub Command9_Click()
    Command9.Visible = False
    Call check_click(Command9)
End Sub

Private Sub Command10_Click()
    Command10.Visible = False
    Call check_click(Command10)
End Sub

Private Sub Command11_Click()
    Command11.Visible = False
    Call check_click(Command11)
End Sub

Private Sub Command12_Click()
    Command12.Visible = False
    Call check_click(Command12)
End Sub

Private Sub Command13_Click()
    Command13.Visible = False
    Call check_click(Command13)
End Sub

Private Sub Command14_Click()
    Command14.Visible = False
    Call check_click(Command14)
End Sub

Private Sub Command15_Click()
    Command15.Visible = False
    Call check_click(Command15)
End Sub

Private Sub Command16_Click()
    Command16.Visible = False
    Call check_click(Command16)
End Sub

Private Sub Command17_Click()
    Command17.Visible = False
    Call check_click(Command17)
End Sub

Private Sub ex_Click()
    If MsgBox("aya kharej mishavid ?", vbYesNo + vbCritical) = vbYes Then
        End
    End If
End Sub

Private Sub Form_Load()
    click = 0
    pic1 = 0
    pic2 = 0
    player = "Player 1"
    Picture3.Picture = Picture1.Picture
    Picture5.Picture = Picture10.Picture
    Picture6.Picture = Picture12.Picture
    Picture8.Picture = Picture2.Picture
    Picture9.Picture = Picture16.Picture
    Picture11.Picture = Picture13.Picture
    Picture14.Picture = Picture7.Picture
    Picture15.Picture = Picture4.Picture
End Sub

Private Sub ng_Click()
    Text1 = "Player 1"
    Text2 = 0
    Text3 = 0
    Text4 = 0
    pic1 = 0
    pic2 = 0
    click = 0
    btn2 = 0
    btn3 = 0
    btn4 = 0
    btn5 = 0
    btn6 = 0
    btn7 = 0
    btn8 = 0
    btn9 = 0
    btn10 = 0
    btn11 = 0
    btn12 = 0
    btn13 = 0
    btn14 = 0
    btn15 = 0
    btn16 = 0
    btn17 = 0
    btn01 = 0
    btn02 = 0
    player = "Player 1"
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    If btn2 <> 1 Then Command2.Visible = True
    If btn3 <> 1 Then Command3.Visible = True
    If btn4 <> 1 Then Command4.Visible = True
    If btn5 <> 1 Then Command5.Visible = True
    If btn6 <> 1 Then Command6.Visible = True
    If btn7 <> 1 Then Command7.Visible = True
    If btn8 <> 1 Then Command8.Visible = True
    If btn9 <> 1 Then Command9.Visible = True
    If btn10 <> 1 Then Command10.Visible = True
    If btn11 <> 1 Then Command11.Visible = True
    If btn12 <> 1 Then Command12.Visible = True
    If btn13 <> 1 Then Command13.Visible = True
    If btn14 <> 1 Then Command14.Visible = True
    If btn15 <> 1 Then Command15.Visible = True
    If btn16 <> 1 Then Command16.Visible = True
    If btn17 <> 1 Then Command17.Visible = True
    Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
    Text2 = Text2 + 1
End Sub
