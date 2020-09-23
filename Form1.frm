VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tasks ToDo..."
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   9495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Warning 
      Caption         =   "!"
      Height          =   315
      Left            =   8880
      TabIndex        =   2
      Top             =   120
      Width           =   435
   End
   Begin VB.TextBox NewTask 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8655
   End
   Begin VB.ListBox Tareas 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2085
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   480
      Width           =   9195
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "SAY NO MORE MSN - VISIT WWW.MeetFindeR.ar.tc and chat with all your friends !!! - SAY NO MORE ICQ "
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   9255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public NoCheckNow As Boolean

Private Sub Form_Load()
On Error GoTo SOL
NoCheckNow = True
Dim Entrada As String

Open App.Path & "\TAREAS.MFR" For Input As #1

    Do Until EOF(1)
    Line Input #1, Entrada
    
        If Mid(Entrada, 1, 5) = "#DONE" Then
            Tareas.AddItem Mid(Entrada, 6), 0
            Tareas.Selected(0) = True
        Else
            Tareas.AddItem Entrada, 0
            
        End If
    Loop
Close #1
NoCheckNow = False
SOL:
    If Err = 53 Then
    
        Open App.Path & "\TAREAS.MFR" For Output As #1: Close #1
        NoCheckNow = False

    End If
End Sub

Private Sub NewTask_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

    If Warning.Value = vbChecked Then
        Tareas.AddItem "!Warning:" & NewTask.Text, 0
    Else
        Tareas.AddItem NewTask.Text, 0
    End If
    
    NewTask = ""
    Warning.Value = vbUnchecked
    NewTask.SetFocus

    Dim i As Integer
    Open App.Path & "\TAREAS.MFR" For Output As #1
    
        For i = 0 To Tareas.ListCount - 1
    
            If Tareas.Selected(i) = True Then
                Print #1, "#DONE" & Tareas.List(i)
            Else
                Print #1, Tareas.List(i)
            End If
    
        Next i
        
    Close #1

End If


End Sub

Private Sub Tareas_ItemCheck(Item As Integer)
    Dim i As Integer
    
If Not NoCheckNow Then
    Open App.Path & "\TAREAS.MFR" For Output As #1
    
        For i = 0 To Tareas.ListCount - 1
    
            If Tareas.Selected(i) = True Then
                Print #1, "#DONE" & Tareas.List(i)
            Else
                Print #1, Tareas.List(i)
            End If
    
        Next i
        
    Close #1
End If
End Sub
