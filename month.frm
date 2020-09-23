VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "My Workaholic PC v2.0"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4020
   FillColor       =   &H00C00000&
   FillStyle       =   0  'Solid
   Icon            =   "month.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   4020
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   7920
      TabIndex        =   28
      Top             =   3840
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   7680
      Top             =   3600
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   7680
      TabIndex        =   25
      Text            =   "0"
      Top             =   3720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text42 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0.00;(0.00)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   23
      Top             =   4080
      Width           =   1335
   End
   Begin VB.TextBox Text31 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0.00;(0.00)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   10
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox Text32 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0.00;(0.00)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   9
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox Text33 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0.00;(0.00)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   8
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox Text34 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0.00;(0.00)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   7
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox Text35 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0.00;(0.00)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   6
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox Text36 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0.00;(0.00)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   5
      Top             =   4080
      Width           =   1335
   End
   Begin VB.TextBox Text37 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0.00;(0.00)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   4
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox Text38 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0.00;(0.00)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   3
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox Text39 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0.00;(0.00)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   2
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox Text40 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0.00;(0.00)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   1
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox Text41 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0.00;(0.00)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   0
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "Breakup Of My Working Time. "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   120
      TabIndex        =   32
      Top             =   1800
      Width           =   3735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Hours Everyday.Below Is The Monthly "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   120
      TabIndex        =   31
      Top             =   1560
      Width           =   3735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   30
      Top             =   1200
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "And I Work For An Average Time Of "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   120
      TabIndex        =   29
      Top             =   840
      Width           =   3735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "This Session I have Been Working For"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   120
      TabIndex        =   27
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label UpLbl 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000012&
      Caption         =   "Seconds"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1680
      MouseIcon       =   "month.frx":0442
      MousePointer    =   99  'Custom
      TabIndex        =   24
      ToolTipText     =   "Click Me"
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label22 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Jan"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label23 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Feb"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label24 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Mar"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label25 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Apr"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "May"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label27 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Jun"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label28 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Jul"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   16
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label29 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Aug"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2160
      TabIndex        =   15
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label30 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Sep"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   14
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label31 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Oct"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   13
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label32 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Nov"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   12
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label33 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Dec"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   11
      Top             =   4080
      Width           =   975
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.ForeColor = &HFFFFFF
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.Show
Unload Me
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.ForeColor = &HFFFFFF
End Sub

Private Sub Label7_Click()
Text1.Text = Val(Text1.Text) + 1
If Val(Text1.Text) = 1 Then
Open App.Path & "\01.cot" For Input As #1
Input #1, jan
Close #1
inmin3 = jan / 60
Text31.Text = inmin3

Open App.Path & "\02.cot" For Input As #1
Input #1, feb
Close #1
inmin3 = feb / 60
Text32.Text = inmin3

Open App.Path & "\03.cot" For Input As #1
Input #1, mar
Close #1
inmin3 = mar / 60
Text33.Text = inmin3

Open App.Path & "\04.cot" For Input As #1
Input #1, ap
Close #1
inmin3 = ap / 60
Text34.Text = inmin3

Open App.Path & "\05.cot" For Input As #1
Input #1, may
Close #1
inmin3 = may / 60
Text35.Text = inmin3

Open App.Path & "\06.cot" For Input As #1
Input #1, jun
Close #1
inmin3 = jun / 60
Text36.Text = inmin3

Open App.Path & "\07.cot" For Input As #1
Input #1, jul
Close #1
inmin3 = jul / 60
Text37.Text = inmin3

Open App.Path & "\08.cot" For Input As #1
Input #1, aug
Close #1
inmin3 = aug / 60
Text38.Text = inmin3

Open App.Path & "\09.cot" For Input As #1
Input #1, sep
Close #1
inmin3 = sep / 60
Text39.Text = inmin3


Open App.Path & "\10.cot" For Input As #1
Input #1, Octo
Close #1
inmin3 = Octo / 60
Text40.Text = inmin3

Open App.Path & "\11.cot" For Input As #1
Input #1, nov
Close #1
inmin3 = nov / 60
Text41.Text = inmin3

Open App.Path & "\12.cot" For Input As #1
Input #1, dec
Close #1
inmin3 = dec
Text42.Text = dec / 60
Text1.Text = Val(Text1.Text) + 1
Label7.Caption = "Minutes"
End If

'Hour

If Val(Text1.Text) = 3 Then
Open App.Path & "\01.cot" For Input As #1
Input #1, jan
Close #1
inmin3 = jan / 3600
Text31.Text = inmin3

Open App.Path & "\02.cot" For Input As #1
Input #1, feb
Close #1
inmin3 = feb / 3600
Text32.Text = inmin3

Open App.Path & "\03.cot" For Input As #1
Input #1, mar
Close #1
inmin3 = mar / 3600
Text33.Text = inmin3

Open App.Path & "\04.cot" For Input As #1
Input #1, ap
Close #1
inmin3 = ap / 3600
Text34.Text = inmin3

Open App.Path & "\05.cot" For Input As #1
Input #1, may
Close #1
inmin3 = may / 3600
Text35.Text = inmin3

Open App.Path & "\06.cot" For Input As #1
Input #1, jun
Close #1
inmin3 = jun / 3600
Text36.Text = inmin3

Open App.Path & "\07.cot" For Input As #1
Input #1, jul
Close #1
inmin3 = jul / 3600
Text37.Text = inmin3

Open App.Path & "\08.cot" For Input As #1
Input #1, aug
Close #1
inmin3 = aug / 3600
Text38.Text = inmin3

Open App.Path & "\09.cot" For Input As #1
Input #1, sep
Close #1
inmin3 = sep / 3600
Text39.Text = inmin3


Open App.Path & "\10.cot" For Input As #1
Input #1, Octo
Close #1
inmin3 = Octo / 3600
Text40.Text = inmin3

Open App.Path & "\11.cot" For Input As #1
Input #1, nov
Close #1
inmin3 = nov / 3600
Text41.Text = inmin3

Open App.Path & "\12.cot" For Input As #1
Input #1, dec
Close #1
inmin3 = dec
Text42.Text = dec / 3600
Text1.Text = Val(Text1.Text) + 1
Label7.Caption = "Hours"
End If

'Days
If Val(Text1.Text) = 5 Then
Open App.Path & "\01.cot" For Input As #1
Input #1, jan
Close #1
inmin3 = jan \ 86400
Text31.Text = inmin3

Open App.Path & "\02.cot" For Input As #1
Input #1, feb
Close #1
inmin3 = feb \ 86400
Text32.Text = inmin3

Open App.Path & "\03.cot" For Input As #1
Input #1, mar
Close #1
inmin3 = mar \ 86400
Text33.Text = inmin3

Open App.Path & "\04.cot" For Input As #1
Input #1, ap
Close #1
inmin3 = ap \ 86400
Text34.Text = inmin3

Open App.Path & "\05.cot" For Input As #1
Input #1, may
Close #1
inmin3 = may \ 86400
Text35.Text = inmin3

Open App.Path & "\06.cot" For Input As #1
Input #1, jun
Close #1
inmin3 = jun \ 86400
Text36.Text = inmin3

Open App.Path & "\07.cot" For Input As #1
Input #1, jul
Close #1
inmin3 = jul \ 86400
Text37.Text = inmin3

Open App.Path & "\08.cot" For Input As #1
Input #1, aug
Close #1
inmin3 = aug \ 86400
Text38.Text = inmin3

Open App.Path & "\09.cot" For Input As #1
Input #1, sep
Close #1
inmin3 = sep \ 86400
Text39.Text = inmin3


Open App.Path & "\10.cot" For Input As #1
Input #1, Octo
Close #1
inmin3 = Octo \ 86400
Text40.Text = inmin3

Open App.Path & "\11.cot" For Input As #1
Input #1, nov
Close #1
inmin3 = nov \ 86400
Text41.Text = inmin3

Open App.Path & "\12.cot" For Input As #1
Input #1, dec
Close #1
inmin3 = dec
Text42.Text = dec \ 86400
Text1.Text = Val(Text1.Text) + 1
Label7.Caption = "Days"
End If


'Seconds


If Val(Text1.Text) = 7 Then
Open App.Path & "\01.cot" For Input As #1
Input #1, jan
Close #1
inmin3 = jan
Text31.Text = inmin3

Open App.Path & "\02.cot" For Input As #1
Input #1, feb
Close #1
inmin3 = feb
Text32.Text = inmin3

Open App.Path & "\03.cot" For Input As #1
Input #1, mar
Close #1
inmin3 = mar
Text33.Text = inmin3

Open App.Path & "\04.cot" For Input As #1
Input #1, ap
Close #1
inmin3 = ap
Text34.Text = inmin3

Open App.Path & "\05.cot" For Input As #1
Input #1, may
Close #1
inmin3 = may
Text35.Text = inmin3

Open App.Path & "\06.cot" For Input As #1
Input #1, jun
Close #1
inmin3 = jun
Text36.Text = inmin3

Open App.Path & "\07.cot" For Input As #1
Input #1, jul
Close #1
inmin3 = jul
Text37.Text = inmin3

Open App.Path & "\08.cot" For Input As #1
Input #1, aug
Close #1
inmin3 = aug
Text38.Text = inmin3

Open App.Path & "\09.cot" For Input As #1
Input #1, sep
Close #1
inmin3 = sep
Text39.Text = inmin3


Open App.Path & "\10.cot" For Input As #1
Input #1, Octo
Close #1
inmin3 = Octo
Text40.Text = inmin3

Open App.Path & "\11.cot" For Input As #1
Input #1, nov
Close #1
inmin3 = nov
Text41.Text = inmin3

Open App.Path & "\12.cot" For Input As #1
Input #1, dec
Close #1
inmin3 = dec
Text42.Text = dec
Text1.Text = 0
Label7.Caption = "Seconds"
End If
End Sub

Private Sub Form_Load()
Open App.Path & "\01.cot" For Input As #1
Input #1, jan
Close #1
inmin3 = jan
Text31.Text = inmin3

Open App.Path & "\02.cot" For Input As #1
Input #1, feb
Close #1
inmin3 = feb
Text32.Text = inmin3

Open App.Path & "\03.cot" For Input As #1
Input #1, mar
Close #1
inmin3 = mar
Text33.Text = inmin3

Open App.Path & "\04.cot" For Input As #1
Input #1, ap
Close #1
inmin3 = ap
Text34.Text = inmin3

Open App.Path & "\05.cot" For Input As #1
Input #1, may
Close #1
inmin3 = may
Text35.Text = inmin3

Open App.Path & "\06.cot" For Input As #1
Input #1, jun
Close #1
inmin3 = jun
Text36.Text = inmin3

Open App.Path & "\07.cot" For Input As #1
Input #1, jul
Close #1
inmin3 = jul
Text37.Text = inmin3

Open App.Path & "\08.cot" For Input As #1
Input #1, aug
Close #1
inmin3 = aug
Text38.Text = inmin3

Open App.Path & "\09.cot" For Input As #1
Input #1, sep
Close #1
inmin3 = sep
Text39.Text = inmin3


Open App.Path & "\10.cot" For Input As #1
Input #1, Octo
Close #1
inmin3 = Octo
Text40.Text = inmin3

Open App.Path & "\11.cot" For Input As #1
Input #1, nov
Close #1
inmin3 = nov
Text41.Text = inmin3

Open App.Path & "\12.cot" For Input As #1
Input #1, dec
Close #1
inmin3 = dec
Text42.Text = dec

'Average
Open App.Path & "\comtime.cot" For Input As #1
Input #1, timea
Close #1

timea1 = timea \ 3600
Open App.Path & "\stdate1.cot" For Input As #1
Input #1, nowdate
Close #1
msg = DateDiff("d", Now, nowdate)
Label3.Caption = timea1 / msg
If msg < 0 Then
msg = DateDiff("d", nowdate, Now)
Label3.Caption = timea1 / (msg + 1)
End If
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.ForeColor = &HFF0000
End Sub

Private Sub Text41_GotFocus()
Command1.SetFocus
End Sub

Private Sub Timer1_Timer()
UpLbl.Caption = FormatCount(GetTickCount, DaysHoursMinutesSecondsMilliseconds)
End Sub

Private Sub UpLbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.ForeColor = &HFFFFFF
End Sub
