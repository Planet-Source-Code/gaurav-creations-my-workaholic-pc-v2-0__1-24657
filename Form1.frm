VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   105
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   285
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   105
   ScaleWidth      =   285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   -240
      TabIndex        =   5
      Text            =   "Text6"
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   -240
      TabIndex        =   4
      Text            =   "Text5"
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   -240
      TabIndex        =   3
      Text            =   "Text4"
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   -240
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Visible         =   0   'False
      Width           =   150
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessID As Long, ByVal dwType As Long) As Long
Private Const RSP_SIMPLE_SERVICE = 1
Private Const RSP_UNREGISTER_SERVICE = 0
Private Sub Form_Load()
Setdefaults:
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN", "comtime", App.Path & "\" & App.EXEName & ".exe"
If App.PrevInstance Then
End
End If
Text2.Text = Now
HideApp
    Open App.Path & "\startup.cot" For Input As #1
    Do While Not EOF(1)
    Input #1, prev
    Loop
    Close #1
    totalr = prev + 1
    
    Open App.Path & "\startup.cot" For Output As #1
  
    Write #1, totalr
    
    Close #1
    
    Open App.Path & "\crash.cot" For Input As #1
    Do While Not EOF(1)
    Input #1, cras
    Loop
    Close #1
    cras1 = cras

    Open App.Path & "\shut.cot" For Input As #1
    Do While Not EOF(1)
    Input #1, prev1
    Loop
    Close #1
    totalr1 = prev1
    Text5.Text = totalr - totalr1
    
    If Val(Text5.Text) > cras1 Then
    Open App.Path & "\crash.cot" For Output As #1
  
    Write #1, Text5.Text
    
    Close #1
    Form2.Show
    End If
    
    Text3.Text = Format$(Now, "mmm d,yyyy")
    Text6.Text = Now
    If totalr = 0 Then
    Open App.Path & "\stdate.cot" For Output As #1
  
    Write #1, Text3.Text
    
    Close #1
    
    Open App.Path & "\stdate1.cot" For Output As #1
  
    Write #1, Text6.Text
    
    Close #1
    
    End If
     
   Text4.Text = Format(Now, "mm")
End Sub

Public Sub HideApp()

Dim process As Long
process = GetCurrentProcessId()
Call RegisterServiceProcess(process, RSP_SIMPLE_SERVICE)

End Sub

Private Sub Form_Unload(Cancel As Integer)
Text1.Text = DateDiff("s", Text2.Text, Now)
If Val(Text1.Text) < 0 Then
Text1.Text = DateDiff("s", Now, Text2.Text)
End If
    
    Open App.Path & "\comtime.cot" For Input As #1
    Do While Not EOF(1)
    Input #1, prevresults
    Loop
    Close #1
    totalresults = prevresults + Val(Text1.Text)
    
    Open App.Path & "\comtime.cot" For Output As #1
  
    Write #1, totalresults
    
    Close #1

    Open App.Path & "\shut.cot" For Input As #1
    Do While Not EOF(1)
    Input #1, prev1
    Loop
    Close #1
    totalr1 = prev1 + 1
    
    Open App.Path & "\shut.cot" For Output As #1
  
    Write #1, totalr1
    
    Close #1

    Open App.Path & "\" & Text4.Text & ".cot" For Input As #1
    Input #1, loadusage2
    Close #1
    inmin3 = loadusage2 + Val(Text1.Text)
       
    Open App.Path & "\" & Text4.Text & ".cot" For Output As #1
  
    Write #1, inmin3
    
    Close #1


End Sub
