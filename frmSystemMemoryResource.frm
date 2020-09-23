VERSION 5.00
Begin VB.Form frmSystemMemoryResource 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System Memory Resource"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   Icon            =   "frmSystemMemoryResource.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
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
   Begin VB.ListBox lstSystemMemoryResource 
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
Attribute VB_Name = "frmSystemMemoryResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGetList_Click()
    On Error Resume Next
    
    Dim SystemMemoryResource As SWbemObject
   
    'Clear current
    lstSystemMemoryResource.Clear
    
    Me.MousePointer = vbHourglass
      
    'Enumerate the instances
    Dim SystemMemoryResourceSet As SWbemObjectSet
    Set SystemMemoryResourceSet = Namespace.InstancesOf("Win32_SystemMemoryResource")
    
    For Each SystemMemoryResource In SystemMemoryResourceSet
        ' Use the RelPath property of the instance path to display the disk
        lstSystemMemoryResource.AddItem SystemMemoryResource.Path_.RelPath
    Next

    Me.MousePointer = vbNormal
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub lstSystemMemoryResource_Click()
    On Error Resume Next
    
    Dim SelectedItem As String
    Dim Value As Variant
    Dim SystemMemoryResource As SWbemObject
    Dim tmpInt As Integer
    
    'Clear current
    lstData.Clear
    
    Me.MousePointer = vbHourglass
    
    SelectedItem = lstSystemMemoryResource.List(lstSystemMemoryResource.ListIndex)
    Set SystemMemoryResource = Namespace.Get(SelectedItem)
    
    Value = SystemMemoryResource.Caption
    lstData.AddItem Left("Caption" & Space(35), 35)
    lstData.List(0) = lstData.List(0) & CStr(Value)

    Value = SystemMemoryResource.CreationClassName
    lstData.AddItem Left("Creation Class Name" & Space(35), 35)
    lstData.List(1) = lstData.List(1) & CStr(Value)

    Value = SystemMemoryResource.CSCreationClassName
    lstData.AddItem Left("CS Creation Class Name" & Space(35), 35)
    lstData.List(2) = lstData.List(2) & CStr(Value)

    Value = SystemMemoryResource.CSName
    lstData.AddItem Left("CS Name" & Space(35), 35)
    lstData.List(3) = lstData.List(3) & CStr(Value)

    Value = SystemMemoryResource.Description
    lstData.AddItem Left("Description" & Space(35), 35)
    lstData.List(4) = lstData.List(4) & CStr(Value)

    Value = SystemMemoryResource.EndingAddress
    lstData.AddItem Left("Ending Address" & Space(35), 35)
    lstData.List(5) = lstData.List(5) & CStr(Value)

    Value = SystemMemoryResource.InstallDate
    lstData.AddItem Left("Install Date" & Space(35), 35)
    lstData.List(6) = lstData.List(6) & CStr(Value)

    Value = SystemMemoryResource.Name
    lstData.AddItem Left("Name" & Space(35), 35)
    lstData.List(7) = lstData.List(7) & CStr(Value)

    Value = SystemMemoryResource.StartingAddress
    lstData.AddItem Left("Starting Address" & Space(35), 35)
    lstData.List(8) = lstData.List(8) & CStr(Value)

    Value = SystemMemoryResource.Status
    lstData.AddItem Left("Status" & Space(35), 35)
    lstData.List(8) = lstData.List(8) & CStr(Value)
    
    Me.MousePointer = vbNormal
End Sub
