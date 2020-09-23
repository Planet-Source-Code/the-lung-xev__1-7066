VERSION 5.00
Begin VB.Form frmPhysicalMemory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Physical Memory"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   Icon            =   "frmPhysicalMemory.frx":0000
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
   Begin VB.ListBox lstPhysicalMemory 
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
Attribute VB_Name = "frmPhysicalMemory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGetList_Click()
    On Error Resume Next
    
    Dim PhysicalMemory As SWbemObject
   
    'Clear current
    lstPhysicalMemory.Clear
    
    Me.MousePointer = vbHourglass
      
    'Enumerate the instances
    Dim PhysicalMemorySet As SWbemObjectSet
    Set PhysicalMemorySet = Namespace.InstancesOf("Win32_PhysicalMemory")
    
    For Each PhysicalMemory In PhysicalMemorySet
        ' Use the RelPath property of the instance path to display the disk
        lstPhysicalMemory.AddItem PhysicalMemory.Path_.RelPath
    Next

    Me.MousePointer = vbNormal
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub lstPhysicalMemory_Click()
    On Error Resume Next
    
    Dim SelectedItem As String
    Dim Value As Variant
    Dim PhysicalMemory As SWbemObject
    Dim tmpInt As Integer
    
    'Clear current
    lstData.Clear
    
    Me.MousePointer = vbHourglass
    
    SelectedItem = lstPhysicalMemory.List(lstPhysicalMemory.ListIndex)
    Set PhysicalMemory = Namespace.Get(SelectedItem)
    
    Value = PhysicalMemory.BankLabel
    lstData.AddItem Left("Bank Label" & Space(35), 35)
    lstData.List(0) = lstData.List(0) & CStr(Value)

    Value = PhysicalMemory.Capacity
    lstData.AddItem Left("Capacity" & Space(35), 35)
    lstData.List(1) = lstData.List(1) & CStr(Value) & "bytes"

    Value = PhysicalMemory.Caption
    lstData.AddItem Left("Caption" & Space(35), 35)
    lstData.List(2) = lstData.List(2) & CStr(Value)

    Value = PhysicalMemory.CreationClassName
    lstData.AddItem Left("Creation Class Name" & Space(35), 35)
    lstData.List(3) = lstData.List(3) & CStr(Value)

    Value = PhysicalMemory.DataWidth
    lstData.AddItem Left("Data Width" & Space(35), 35)
    lstData.List(4) = lstData.List(4) & CStr(Value) & "bits"

    Value = PhysicalMemory.Description
    lstData.AddItem Left("Description" & Space(35), 35)
    lstData.List(5) = lstData.List(5) & CStr(Value)

    Value = PhysicalMemory.DeviceLocator
    lstData.AddItem Left("Device Locator" & Space(35), 35)
    lstData.List(6) = lstData.List(6) & CStr(Value)

    Value = PhysicalMemory.FormFactor
    lstData.AddItem Left("Form Factor" & Space(35), 35)
    Select Case Value
        Case 0: lstData.List(7) = lstData.List(7) & "Unknown"
        Case 1: lstData.List(7) = lstData.List(7) & "Other"
        Case 2: lstData.List(7) = lstData.List(7) & "SIP"
        Case 3: lstData.List(7) = lstData.List(7) & "DIP"
        Case 4: lstData.List(7) = lstData.List(7) & "ZIP"
        Case 5: lstData.List(7) = lstData.List(7) & "SOJ"
        Case 6: lstData.List(7) = lstData.List(7) & "Proprietary"
        Case 7: lstData.List(7) = lstData.List(7) & "SIMM"
        Case 8: lstData.List(7) = lstData.List(7) & "DIMM"
        Case 9: lstData.List(7) = lstData.List(7) & "TSOP"
        Case 10: lstData.List(7) = lstData.List(7) & "PGA"
        Case 11: lstData.List(7) = lstData.List(7) & "RIMM"
        Case 12: lstData.List(7) = lstData.List(7) & "SODIMM"
    End Select

    Value = PhysicalMemory.HotSwappable
    lstData.AddItem Left("Hot Swappable" & Space(35), 35)
    lstData.List(8) = lstData.List(8) & CStr(Value)

    Value = PhysicalMemory.InstallDate
    lstData.AddItem Left("Install Date" & Space(35), 35)
    lstData.List(9) = lstData.List(9) & CStr(Value)

    Value = PhysicalMemory.InterleaveDataDepth
    lstData.AddItem Left("Interleave Data Depth" & Space(35), 35)
    lstData.List(10) = lstData.List(10) & CStr(Value)

    Value = PhysicalMemory.InterleavePosition
    lstData.AddItem Left("Interleave Position" & Space(35), 35)
    Select Case Value
        Case 0: lstData.List(11) = lstData.List(11) & "non-interleaved"
        Case 1: lstData.List(11) = lstData.List(11) & "first position"
        Case 2: lstData.List(11) = lstData.List(11) & "second position"
    End Select

    Value = PhysicalMemory.Manufacturer
    lstData.AddItem Left("Manufacturer" & Space(35), 35)
    lstData.List(12) = lstData.List(12) & CStr(Value)

    Value = PhysicalMemory.MemoryType
    lstData.AddItem Left("Memory Type" & Space(35), 35)
    Select Case Value
        Case 1: lstData.List(13) = lstData.List(13) & "Unknown"
        Case 2: lstData.List(13) = lstData.List(13) & "Other"
        Case 3: lstData.List(13) = lstData.List(13) & "DRAM"
        Case 4: lstData.List(13) = lstData.List(13) & "Synchronous DRAM"
        Case 5: lstData.List(13) = lstData.List(13) & "Cache DRAM"
        Case 6: lstData.List(13) = lstData.List(13) & "EDO"
        Case 7: lstData.List(13) = lstData.List(13) & "EDRAM"
        Case 8: lstData.List(13) = lstData.List(13) & "VRAM"
        Case 9: lstData.List(13) = lstData.List(13) & "SRAM"
        Case 10: lstData.List(13) = lstData.List(13) & "RAM"
        Case 11: lstData.List(13) = lstData.List(13) & "ROM"
        Case 12: lstData.List(13) = lstData.List(13) & "Flash"
        Case 13: lstData.List(13) = lstData.List(13) & "EEPROM"
        Case 14: lstData.List(13) = lstData.List(13) & "FEPROM"
        Case 15: lstData.List(13) = lstData.List(13) & "EPROM"
        Case 16: lstData.List(13) = lstData.List(13) & "CDRAM"
        Case 17: lstData.List(13) = lstData.List(13) & "3DRAM"
        Case 18: lstData.List(13) = lstData.List(13) & "SDRAM"
        Case 19: lstData.List(13) = lstData.List(13) & "SGRAM"
    End Select

    Value = PhysicalMemory.Model
    lstData.AddItem Left("Model" & Space(35), 35)
    lstData.List(14) = lstData.List(14) & CStr(Value)

    Value = PhysicalMemory.Name
    lstData.AddItem Left("Name" & Space(35), 35)
    lstData.List(15) = lstData.List(15) & CStr(Value)

    Value = PhysicalMemory.OtherIdentifyingInfo
    lstData.AddItem Left("Other Identifying Info" & Space(35), 35)
    lstData.List(16) = lstData.List(16) & CStr(Value)

    Value = PhysicalMemory.PartNumber
    lstData.AddItem Left("Part Number" & Space(35), 35)
    lstData.List(17) = lstData.List(17) & CStr(Value)

    Value = PhysicalMemory.PositionInRow
    lstData.AddItem Left("Position In Row" & Space(35), 35)
    lstData.List(18) = lstData.List(18) & CStr(Value)

    Value = PhysicalMemory.PoweredOn
    lstData.AddItem Left("Powered On" & Space(35), 35)
    lstData.List(19) = lstData.List(19) & CStr(Value)

    Value = PhysicalMemory.Removable
    lstData.AddItem Left("Removable" & Space(35), 35)
    lstData.List(20) = lstData.List(20) & CStr(Value)

    Value = PhysicalMemory.Replaceable
    lstData.AddItem Left("Replaceable" & Space(35), 35)
    lstData.List(21) = lstData.List(21) & CStr(Value)

    Value = PhysicalMemory.SerialNumber
    lstData.AddItem Left("Serial Number" & Space(35), 35)
    lstData.List(22) = lstData.List(22) & CStr(Value)

    Value = PhysicalMemory.SKU
    lstData.AddItem Left("SKU" & Space(35), 35)
    lstData.List(23) = lstData.List(23) & CStr(Value)

    Value = PhysicalMemory.Speed
    lstData.AddItem Left("Speed" & Space(35), 35)
    lstData.List(24) = lstData.List(24) & CStr(Value) & "nanoseconds"

    Value = PhysicalMemory.Status
    lstData.AddItem Left("Status" & Space(35), 35)
    lstData.List(25) = lstData.List(25) & CStr(Value)

    Value = PhysicalMemory.Tag
    lstData.AddItem Left("Tag" & Space(35), 35)
    lstData.List(26) = lstData.List(26) & CStr(Value)

    Value = PhysicalMemory.TotalWidth
    lstData.AddItem Left("Total Width" & Space(35), 35)
    lstData.List(27) = lstData.List(27) & CStr(Value) & "bits"

    Value = PhysicalMemory.TypeDetail
    lstData.AddItem Left("Type Detail" & Space(35), 35)
    Select Case Value
        Case 1: lstData.List(28) = lstData.List(28) & "Reserved"
        Case 1: lstData.List(28) = lstData.List(28) & "Other"
        Case 1: lstData.List(28) = lstData.List(28) & "Unknown"
        Case 1: lstData.List(28) = lstData.List(28) & "Fast-paged"
        Case 1: lstData.List(28) = lstData.List(28) & "Static column"
        Case 1: lstData.List(28) = lstData.List(28) & "Pseudo-static"
        Case 1: lstData.List(28) = lstData.List(28) & "RAMBUS"
        Case 1: lstData.List(28) = lstData.List(28) & "Synchronous"
        Case 1: lstData.List(28) = lstData.List(28) & "CMOS"
        Case 1: lstData.List(28) = lstData.List(28) & "EDO"
        Case 1: lstData.List(28) = lstData.List(28) & "Window DRAM"
        Case 1: lstData.List(28) = lstData.List(28) & "Cache DRAM"
        Case 1: lstData.List(28) = lstData.List(28) & "Non-volatile"
    End Select

    Value = PhysicalMemory.Version
    lstData.AddItem Left("Version" & Space(35), 35)
    lstData.List(29) = lstData.List(29) & CStr(Value)
    
    Me.MousePointer = vbNormal
End Sub
