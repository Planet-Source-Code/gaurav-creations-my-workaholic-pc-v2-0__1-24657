VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H80000007&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "My Workaholic PC -Time Adder"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6180
   Icon            =   "crash.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   4680
      TabIndex        =   6
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4560
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Please Confirm And Then Click Me"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   2160
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Left            =   2520
      MaxLength       =   4
      TabIndex        =   1
      Text            =   "0"
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000007&
      Caption         =   "Please Note : You Wont't Be Able To Edit Afterwards"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   1800
      Width           =   6255
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      Caption         =   "Enter In Minutes Only."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   $"crash.frx":0442
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text2.Text = Val(Text1.Text) * 60
    Open App.Path & "\comtime.cot" For Input As #1
    Do While Not EOF(1)
    Input #1, prevresults
    Loop
    Close #1
    totalresults = prevresults + Val(Text2.Text)
    
    Open App.Path & "\comtime.cot" For Output As #1
  
    Write #1, totalresults
    
    Close #1
    
    Open App.Path & "\" & Text4.Text & ".cot" For Input As #1
    Input #1, loadusage2
    Close #1
    inmin3 = loadusage2 + Val(Text2.Text)
       
    Open App.Path & "\" & Text4.Text & ".cot" For Output As #1
  
    Write #1, inmin3
    
    Close #1
Unload Me
End Sub

Private Sub Form_Load()
Text4.Text = Format(Now, "mm")
End Sub

Private Sub Text1_Change()
If Not IsNumeric(Text1.Text) Or Val(Text1.Text) < 0 Or Val(Text1.Text) > 9999 Then
Text1.Text = 0
msg = MsgBox("Enter Value Between 0 and 9999 Only", vbCritical, "My Workaholic PC")
End If
End Sub
