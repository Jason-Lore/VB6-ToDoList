VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ToDo 
   Caption         =   "To Do List"
   ClientHeight    =   5775
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   8550
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   3375
      Left            =   720
      TabIndex        =   3
      Top             =   1440
      Width           =   7215
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   5160
   End
   Begin MSComCtl2.DTPicker dt_input 
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   240
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      Format          =   2752513
      CurrentDate     =   45553
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Task"
      Height          =   495
      Left            =   3600
      TabIndex        =   1
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox txt_input 
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Text            =   "Enter your task here."
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "ToDo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If txt_input.Text = "" Or txt_input.Text = "Enter your task here." Then
        MsgBox ("Enter in a task item")
    Else
        AddTask
    End If
End Sub

Private Sub AddTask()
    Dim taskDescription As String
    Dim taskDate As String
    
    taskDescription = txt_input.Text
    taskDate = dt_input.Value
    List1.AddItem taskDescription & " - " & Format(taskDate, "mm/dd/yyyy")

End Sub

Private Sub HighlightPastDueTasks()
    Dim i As Integer
    Dim taskDate As Date
    Dim taskDescription As String
    Dim dateStr As String
    Dim pos As Integer
    Dim suffix As String
    
    suffix = " - Overdue"
    
        For i = 0 To List1.ListCount - 1
            taskDescription = List1.List(i)
        
            pos = InStr(taskDescription, "-") + 2
        
            dateStr = Trim(Mid(taskDescription, pos, 10))
        
            On Error Resume Next
            taskDate = CDate(dateStr)
            If Err.Number <> 0 Then
                Err.Clear
                GoTo ContinueLoop
            End If
            On Error GoTo 0
        
            If taskDate < Date And Right(List1.List(i), Len(suffix)) <> suffix Then
                List1.List(i) = List1.List(i) & " - Overdue"
            End If
ContinueLoop:
    Next i
End Sub

Private Sub Form_Load()
Dim fileNum As Integer
Dim filePath As String
Dim lineText As String

    filePath = "C:\Program Files\ToDo\tasks.txt"
    fileNum = FreeFile
    
    If Dir(filePath) <> "" Then
        Open filePath For Input As #fileNum
    
        While Not EOF(fileNum)
            Line Input #fileNum, lineText
            List1.AddItem lineText
        Wend
        Close #fileNum
    End If
    Timer1.Interval = 1000
    Timer1.Enabled = True
End Sub

Private Sub List1_dblClick()
Dim response As Integer

response = MsgBox("Are you sure you want to mark this item as complete?" & vbCrLf & "(This will remove it from your view)", vbYesNo + vbQuestion, "Mark as complete")

    If response = vbYes Then
        List1.RemoveItem List1.ListIndex
    End If
    
End Sub

Private Sub Timer1_Timer()
    HighlightPastDueTasks
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    Dim fileNum As Integer
    Dim filePath As String
    
    filePath = "C:\Program Files\ToDo\tasks.txt"
    fileNum = FreeFile
    
    Open filePath For Output As #fileNum
    
    For i = 0 To List1.ListCount - 1
        Print #fileNum, List1.List(i)
    Next i
    
    Close #fileNum
    
End Sub


