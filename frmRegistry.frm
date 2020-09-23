VERSION 5.00
Begin VB.Form frmRegistry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registry"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   Icon            =   "frmRegistry.frx":0000
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
   Begin VB.ListBox lstRegistry 
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
Attribute VB_Name = "frmRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGetList_Click()
    On Error Resume Next
    
    Dim Registry As SWbemObject
   
    'Clear current
    lstRegistry.Clear
    
    Me.MousePointer = vbHourglass
      
    'Enumerate the instances
    Dim RegistrySet As SWbemObjectSet
    Set RegistrySet = Namespace.InstancesOf("Win32_Registry")
    
    For Each Registry In RegistrySet
        ' Use the RelPath property of the instance path to display the disk
        lstRegistry.AddItem Registry.Path_.RelPath
    Next

    Me.MousePointer = vbNormal
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub lstRegistry_Click()
    On Error Resume Next
    
    Dim SelectedItem As String
    Dim Value As Variant
    Dim Registry As SWbemObject
    Dim tmpInt As Integer

    'Clear current
    lstData.Clear
    
    Me.MousePointer = vbHourglass

    SelectedItem = lstRegistry.List(lstRegistry.ListIndex)
    Set Registry = Namespace.Get(SelectedItem)
    
    Value = Registry.Caption
    lstData.AddItem Left("Caption" & Space(35), 35)
    lstData.List(0) = lstData.List(0) & CStr(Value)

    Value = Registry.CurrentSize
    lstData.AddItem Left("Current Size" & Space(35), 35)
    lstData.List(1) = lstData.List(1) & CStr(Value)

    Value = Registry.Description
    lstData.AddItem Left("Description" & Space(35), 35)
    lstData.List(2) = lstData.List(2) & CStr(Value)

    Value = Registry.InstallDate
    lstData.AddItem Left("Install Date" & Space(35), 35)
    lstData.List(3) = lstData.List(3) & CStr(Value)

    Value = Registry.MaximumSize
    lstData.AddItem Left("Maximum Size" & Space(35), 35)
    lstData.List(4) = lstData.List(4) & CStr(Value) & "Megabytes"

    Value = Registry.Name
    lstData.AddItem Left("Name" & Space(35), 35)
    lstData.List(5) = lstData.List(5) & CStr(Value)

    Value = Registry.ProposedSize
    lstData.AddItem Left("Proposed Size" & Space(35), 35)
    lstData.List(6) = lstData.List(6) & CStr(Value) & "Megabytes"

    Value = Registry.Status
    lstData.AddItem Left("Status" & Space(35), 35)
    lstData.List(7) = lstData.List(7) & CStr(Value)
    
    Me.MousePointer = vbNormal
End Sub
