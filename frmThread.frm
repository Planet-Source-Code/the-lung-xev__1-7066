VERSION 5.00
Begin VB.Form frmThread 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Thread"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   Icon            =   "frmThread.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstData 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2205
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   8655
   End
   Begin VB.ListBox lstThread 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   8655
   End
   Begin VB.CommandButton cmdGetList 
      Caption         =   "Get List"
      Height          =   350
      Left            =   7800
      TabIndex        =   4
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label lblData 
      Caption         =   "Data"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label lblList 
      Caption         =   "List"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmThread"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGetList_Click()
    On Error Resume Next
    
    Dim Thread As SWbemObject
   
    'Clear current
    lstThread.Clear
    
    Me.MousePointer = vbHourglass
      
    'Enumerate the instances
    Dim ThreadSet As SWbemObjectSet
    Set ThreadSet = Namespace.InstancesOf("Win32_Thread")
    
    For Each Thread In ThreadSet
        ' Use the RelPath property of the instance path to display the disk
        lstThread.AddItem Thread.Path_.RelPath
    Next

    Me.MousePointer = vbNormal
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub lstThread_Click()
    On Error Resume Next
    
    Dim SelectedItem As String
    Dim Value As Variant
    Dim Thread As SWbemObject
    Dim tmpInt As Integer
    
    'Clear current
    lstData.Clear

    Me.MousePointer = vbHourglass
    
    SelectedItem = lstThread.List(lstThread.ListIndex)
    Set Thread = Namespace.Get(SelectedItem)
    
    Value = Thread.Caption
    lstData.AddItem Left("Caption" & Space(35), 35)
    lstData.List(0) = lstData.List(0) & CStr(Value)

    Value = Thread.CreationClassName
    lstData.AddItem Left("Creation Class Name" & Space(35), 35)
    lstData.List(1) = lstData.List(1) & CStr(Value)

    Value = Thread.CSCreationClassName
    lstData.AddItem Left("CS Creation Class Name" & Space(35), 35)
    lstData.List(2) = lstData.List(2) & CStr(Value)
    
    Value = Thread.CSName
    lstData.AddItem Left("CS Name" & Space(35), 35)
    lstData.List(3) = lstData.List(3) & CStr(Value)

    Value = Thread.Description
    lstData.AddItem Left("Description" & Space(35), 35)
    lstData.List(4) = lstData.List(4) & CStr(Value)

    Value = Thread.ElapsedTime
    lstData.AddItem Left("Elapsed Time" & Space(35), 35)
    lstData.List(5) = lstData.List(5) & CStr(Value)

    Value = Thread.ExecutionState
    lstData.AddItem Left("ExecutionState" & Space(35), 35)
    Select Case Value
        Case 0: lstData.List(6) = lstData.List(6) & "Unknown"
        Case 1: lstData.List(6) = lstData.List(6) & "Other"
        Case 2: lstData.List(6) = lstData.List(6) & "Ready"
        Case 3: lstData.List(6) = lstData.List(6) & "Running"
        Case 4: lstData.List(6) = lstData.List(6) & "Blocked"
        Case 5: lstData.List(6) = lstData.List(6) & "Suspended Blocked"
        Case 6: lstData.List(6) = lstData.List(6) & "Suspended Ready"
    End Select

    Value = Thread.Handle
    lstData.AddItem Left("Handle" & Space(35), 35)
    lstData.List(7) = lstData.List(7) & CStr(Value)

    Value = Thread.InstallDate
    lstData.AddItem Left("Install Date" & Space(35), 35)
    lstData.List(8) = lstData.List(8) & CStr(Value)

    Value = Thread.KernelModeTime
    lstData.AddItem Left("Kernel Mode Time" & Space(35), 35)
    lstData.List(9) = lstData.List(9) & CStr(Value) & "milliseconds"

    Value = Thread.Name
    lstData.AddItem Left("Name" & Space(35), 35)
    lstData.List(10) = lstData.List(10) & CStr(Value)

    Value = Thread.OSCreationClassName
    lstData.AddItem Left("OS Creation Class Name" & Space(35), 35)
    lstData.List(11) = lstData.List(11) & CStr(Value)

    Value = Thread.OSName
    lstData.AddItem Left("OS Name" & Space(35), 35)
    lstData.List(12) = lstData.List(12) & CStr(Value)

    Value = Thread.Priority
    lstData.AddItem Left("Priority" & Space(35), 35)
    lstData.List(13) = lstData.List(13) & CStr(Value)

    Value = Thread.PriorityBase
    lstData.AddItem Left("Priority Base" & Space(35), 35)
    lstData.List(14) = lstData.List(14) & CStr(Value)

    Value = Thread.ProcessCreationClassName
    lstData.AddItem Left("Process Creation Class Name" & Space(35), 35)
    lstData.List(15) = lstData.List(15) & CStr(Value)

    Value = Thread.ProcessHandle
    lstData.AddItem Left("Process Handle" & Space(35), 35)
    lstData.List(16) = lstData.List(16) & CStr(Value)

    Value = Thread.StartAddress
    lstData.AddItem Left("Start Address" & Space(35), 35)
    lstData.List(17) = lstData.List(17) & CStr(Value)

    Value = Thread.Status
    lstData.AddItem Left("Status" & Space(35), 35)
    lstData.List(18) = lstData.List(18) & CStr(Value)

    Value = Thread.ThreadState
    lstData.AddItem Left("Thread State" & Space(35), 35)
    Select Case Value
        Case 1: lstData.List(19) = lstData.List(19) & "Initialized (recognized by the microkernel)"
        Case 1: lstData.List(19) = lstData.List(19) & "Ready (prepared to run on next available processor)"
        Case 1: lstData.List(19) = lstData.List(19) & "Running (executing)"
        Case 1: lstData.List(19) = lstData.List(19) & "Standby (about to run, only one thread may be in this state at a time)"
        Case 1: lstData.List(19) = lstData.List(19) & "Terminated (finished executing)"
        Case 1: lstData.List(19) = lstData.List(19) & "Waiting (not ready for the processor, when ready, it will be rescheduled)"
        Case 1: lstData.List(19) = lstData.List(19) & "Transition (thread is waiting for resources other than the processor)"
        Case 1: lstData.List(19) = lstData.List(19) & "Unknown (thread state is unknown)"
    End Select

    Value = Thread.ThreadWaitReason
    lstData.AddItem Left("Thread Wait Reason" & Space(35), 35)
    Select Case Value
        Case 0: lstData.List(20) = lstData.List(20) & "Executive"
        Case 1: lstData.List(20) = lstData.List(20) & "FreePage"
        Case 2: lstData.List(20) = lstData.List(20) & "PageIn"
        Case 3: lstData.List(20) = lstData.List(20) & "PoolAllocation"
        Case 4: lstData.List(20) = lstData.List(20) & "ExecutionDelay"
        Case 5: lstData.List(20) = lstData.List(20) & "FreePage"
        Case 6: lstData.List(20) = lstData.List(20) & "PageIn"
        Case 7: lstData.List(20) = lstData.List(20) & "Executive"
        Case 8: lstData.List(20) = lstData.List(20) & "FreePage"
        Case 9: lstData.List(20) = lstData.List(20) & "PageIn"
        Case 10: lstData.List(20) = lstData.List(20) & "PoolAllocation"
        Case 11: lstData.List(20) = lstData.List(20) & "ExecutionDelay"
        Case 12: lstData.List(20) = lstData.List(20) & "FreePage"
        Case 13: lstData.List(20) = lstData.List(20) & "PageIn"
        Case 14: lstData.List(20) = lstData.List(20) & "EventPairHigh"
        Case 15: lstData.List(20) = lstData.List(20) & "EventPairLow"
        Case 16: lstData.List(20) = lstData.List(20) & "LPCReceive"
        Case 17: lstData.List(20) = lstData.List(20) & "LPCReply"
        Case 18: lstData.List(20) = lstData.List(20) & "VirtualMemory"
        Case 19: lstData.List(20) = lstData.List(20) & "PageOut"
        Case 20: lstData.List(20) = lstData.List(20) & "Unknown"
    End Select

    Value = Thread.UserModeTime
    lstData.AddItem Left("Status" & Space(35), 35)
    lstData.List(21) = lstData.List(21) & CStr(Value) & "milliseconds"
    
    Me.MousePointer = vbNormal
End Sub
