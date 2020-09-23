VERSION 5.00
Begin VB.Form frmCacheMemory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cache Memory"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   Icon            =   "frmCacheMemory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
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
   Begin VB.ListBox lstCurrentSRAM 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   120
      TabIndex        =   5
      Top             =   4080
      Width           =   4215
   End
   Begin VB.ListBox lstPowerManagementCapabilities 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   4560
      TabIndex        =   7
      Top             =   4080
      Width           =   4215
   End
   Begin VB.ListBox lstCache 
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
      TabIndex        =   8
      Top             =   4680
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
   Begin VB.Label lblCurrentSRAM 
      Caption         =   "Current SRAM"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Label lblPowerManagementCapabilities 
      Caption         =   "Power Management Capabilities"
      Height          =   255
      Left            =   4560
      TabIndex        =   6
      Top             =   3840
      Width           =   2295
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
Attribute VB_Name = "frmCacheMemory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGetList_Click()
    On Error Resume Next
    
    Dim Cache As SWbemObject
   
    'Clear current
    lstCache.Clear
    
    Me.MousePointer = vbHourglass
      
    'Enumerate the instances
    Dim CacheSet As SWbemObjectSet
    Set CacheSet = Namespace.InstancesOf("Win32_CacheMemory")
    
    For Each Cache In CacheSet
        ' Use the RelPath property of the instance path to display the disk
        lstCache.AddItem Cache.Path_.RelPath
    Next

    Me.MousePointer = vbNormal
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub lstCache_Click()
    On Error Resume Next
    
    Dim SelectedItem As String
    Dim Value As Variant
    Dim Cache As SWbemObject
    Dim tmpInt As Integer
    
    'Clear current
    lstData.Clear
    lstCurrentSRAM.Clear
    lstPowerManagementCapabilities.Clear

    Me.MousePointer = vbHourglass
    
    SelectedItem = lstCache.List(lstCache.ListIndex)
    Set Cache = Namespace.Get(SelectedItem)
    
    Value = Cache.Availability
    lstData.AddItem Left("Availability" & Space(35), 35)
    Select Case Value
        Case 1: lstData.List(0) = lstData.List(0) & "Other"
        Case 2: lstData.List(0) = lstData.List(0) & "Unknown"
        Case 3: lstData.List(0) = lstData.List(0) & "Running/Full Power"
        Case 4: lstData.List(0) = lstData.List(0) & "Warning"
        Case 5: lstData.List(0) = lstData.List(0) & "In Test"
        Case 6: lstData.List(0) = lstData.List(0) & "Not Applicable"
        Case 7: lstData.List(0) = lstData.List(0) & "Power Off"
        Case 8: lstData.List(0) = lstData.List(0) & "Off Line"
        Case 9: lstData.List(0) = lstData.List(0) & "Off Duty"
        Case 10: lstData.List(0) = lstData.List(0) & "Degraded"
        Case 11: lstData.List(0) = lstData.List(0) & "Not Installed"
        Case 12: lstData.List(0) = lstData.List(0) & "Install Error"
        Case 13: lstData.List(0) = lstData.List(0) & "Power Save - Unknown"
        Case 14: lstData.List(0) = lstData.List(0) & "Power Save - Low Power Mode"
        Case 15: lstData.List(0) = lstData.List(0) & "Power Save - Standby"
        Case 16: lstData.List(0) = lstData.List(0) & "Power Cycle"
        Case 17: lstData.List(0) = lstData.List(0) & "Power Save - Warning"
    End Select
    
    Value = Cache.Access
    lstData.AddItem Left("Access" & Space(35), 35)
    Select Case Value
        Case 0: lstData.List(1) = lstData.List(1) & "Unknown"
        Case 1: lstData.List(1) = lstData.List(1) & "Readable"
        Case 2: lstData.List(1) = lstData.List(1) & "Writable"
        Case 3: lstData.List(1) = lstData.List(1) & "Read/Write Supported"
    End Select
    
    Value = Cache.AdditionalErrorData
    lstData.AddItem Left("Additional Error Data" & Space(35), 35)
    lstData.List(2) = lstData.List(2) & CStr(Value)
    
    Value = Cache.Associativity
    lstData.AddItem Left("Associativity" & Space(35), 35)
    Select Case Value
        Case 1: lstData.List(3) = lstData.List(3) & "Other"
        Case 2: lstData.List(3) = lstData.List(3) & "Unknown"
        Case 3: lstData.List(3) = lstData.List(3) & "Direct Mapped"
        Case 4: lstData.List(3) = lstData.List(3) & "2-way Set-Associative"
        Case 5: lstData.List(3) = lstData.List(3) & "4-way Set-Associative"
        Case 6: lstData.List(3) = lstData.List(3) & "Fully Associative"
    End Select
    
    Value = Cache.BlockSize
    lstData.AddItem Left("BlockSize" & Space(35), 35)
    lstData.List(4) = lstData.List(4) & CStr(Value) & "bytes"
    
    Value = Cache.CacheSpeed
    lstData.AddItem Left("Cache Speed" & Space(35), 35)
    lstData.List(5) = lstData.List(5) & CStr(Value) & "NanoSeconds"
    
    Value = Cache.CacheType
    lstData.AddItem Left("Cache Type" & Space(35), 35)
    Select Case Value
        Case 1: lstData.List(6) = lstData.List(6) & "Other"
        Case 2: lstData.List(6) = lstData.List(6) & "Unknown"
        Case 3: lstData.List(6) = lstData.List(6) & "Instruction"
        Case 4: lstData.List(6) = lstData.List(6) & "Data"
        Case 5: lstData.List(6) = lstData.List(6) & "Unified"
    End Select
    
    Value = Cache.Caption
    lstData.AddItem Left("Caption" & Space(35), 35)
    lstData.List(7) = lstData.List(7) & CStr(Value)
    
    Value = Cache.ConfigManagerErrorCode
    lstData.AddItem Left("Config Manager Error Code" & Space(35), 35)
    Select Case Value
        Case 0: lstData.List(8) = lstData.List(8) & "This device is working properly."
        Case 1: lstData.List(8) = lstData.List(8) & "This device is not configured correctly."
        Case 2: lstData.List(8) = lstData.List(8) & "Windows cannot load the driver for this device."
        Case 3: lstData.List(8) = lstData.List(8) & "The driver for this device might be corrupted, or your system may be running low on memory or other resources."
        Case 4: lstData.List(8) = lstData.List(8) & "This device is not working properly. One of its drivers or your registry might be corrupted."
        Case 5: lstData.List(8) = lstData.List(8) & "The driver for this device needs a resource that Windows cannot manage."
        Case 6: lstData.List(8) = lstData.List(8) & "The boot configuration for this device conflicts with other devices."
        Case 7: lstData.List(8) = lstData.List(8) & "Cannot filter."
        Case 8: lstData.List(8) = lstData.List(8) & "The driver loader for the device is missing."
        Case 9: lstData.List(8) = lstData.List(8) & "This device is not working properly because the controlling firmware is reporting the resources for the device incorrectly."
        Case 10: lstData.List(8) = lstData.List(8) & "This device cannot start."
        Case 11: lstData.List(8) = lstData.List(8) & "This device failed."
        Case 12: lstData.List(8) = lstData.List(8) & "This device cannot find enough free resources that it can use."
        Case 13: lstData.List(8) = lstData.List(8) & "Windows cannot verify this device's resources."
        Case 14: lstData.List(8) = lstData.List(8) & "This device cannot work properly until you restart your computer."
        Case 15: lstData.List(8) = lstData.List(8) & "This device is not working properly because there is probably a re-enumeration problem. "
        Case 16: lstData.List(8) = lstData.List(8) & "Windows cannot identify all the resources this device uses."
        Case 17: lstData.List(8) = lstData.List(8) & "This device is asking for an unknown resource type."
        Case 18: lstData.List(8) = lstData.List(8) & "Reinstall the drivers for this device."
        Case 19: lstData.List(8) = lstData.List(8) & "Your registry might be corrupted."
        Case 20: lstData.List(8) = lstData.List(8) & "System failure: Try changing the driver for this device. If that does not work, see your hardware documentation."
        Case 21: lstData.List(8) = lstData.List(8) & "Windows is removing this device."
        Case 22: lstData.List(8) = lstData.List(8) & "This device is disabled."
        Case 23: lstData.List(8) = lstData.List(8) & "System failure: Try changing the driver for this device. If that doesn't work, see your hardware documentation."
        Case 24: lstData.List(8) = lstData.List(8) & "This device is not present, is not working properly, or does not have all its drivers installed."
        Case 25: lstData.List(8) = lstData.List(8) & "Windows is still setting up this device."
        Case 26: lstData.List(8) = lstData.List(8) & "Windows is still setting up this device."
        Case 27: lstData.List(8) = lstData.List(8) & "This device does not have valid log configuration."
        Case 28: lstData.List(8) = lstData.List(8) & "The drivers for this device are not installed."
        Case 29: lstData.List(8) = lstData.List(8) & "This device is disabled because the firmware of the device did not give it the required resources."
        Case 30: lstData.List(8) = lstData.List(8) & "This device is using an Interrupt Request (IRQ) resource that another device is using."
        Case 31: lstData.List(8) = lstData.List(8) & "This device is not working properly because Windows cannot load the drivers required for this device."
    End Select
    
    Value = Cache.ConfigManagerUserConfig
    lstData.AddItem Left("Config Manager User Config" & Space(35), 35)
    lstData.List(9) = lstData.List(9) & CStr(Value)
    
    Value = Cache.CorrectableError
    lstData.AddItem Left("Correctable Error" & Space(35), 35)
    lstData.List(10) = lstData.List(10) & CStr(Value)
    
    Value = Cache.CreationClassName
    lstData.AddItem Left("Creation Class Name" & Space(35), 35)
    lstData.List(11) = lstData.List(11) & CStr(Value)
    
    Value = Cache.CurrentSRAM
    For tmpInt = 1 To 7
        Select Case Value(tmpInt) 'Put in data
            Case 1: lstCurrentSRAM.AddItem "Other"
            Case 2: lstCurrentSRAM.AddItem "Unknown"
            Case 3: lstCurrentSRAM.AddItem "Non-Burst"
            Case 4: lstCurrentSRAM.AddItem "Burst"
            Case 5: lstCurrentSRAM.AddItem "Pipeline Burst"
            Case 6: lstCurrentSRAM.AddItem "Synchronous"
            Case 7: lstCurrentSRAM.AddItem "Asynchronous"
            Case Else: Exit For
        End Select
    Next tmpInt
    
    Value = Cache.Description
    lstData.AddItem Left("Description" & Space(35), 35)
    lstData.List(12) = lstData.List(12) & CStr(Value)
    
    Value = Cache.DeviceID
    lstData.AddItem Left("Device ID" & Space(35), 35)
    lstData.List(13) = lstData.List(13) & CStr(Value)
    
    Value = Cache.EndingAddress
    lstData.AddItem Left("Ending Address" & Space(35), 35)
    lstData.List(14) = lstData.List(14) & CStr(Value) & "kilobytes"
    
    Value = Cache.ErrorAccess
    lstData.AddItem Left("Error Access" & Space(35), 35)
    Select Case Value
        Case 1: lstData.List(15) = lstData.List(15) & "Other"
        Case 2: lstData.List(15) = lstData.List(15) & "Unknown"
        Case 3: lstData.List(15) = lstData.List(15) & "Read"
        Case 4: lstData.List(15) = lstData.List(15) & "Write"
        Case 5: lstData.List(15) = lstData.List(15) & "Partial Write"
    End Select
    
    Value = Cache.ErrorAddres
    lstData.AddItem Left("Error Addres" & Space(35), 35)
    lstData.List(16) = lstData.List(16) & CStr(Value)
    
    Value = Cache.ErrorCleared
    lstData.AddItem Left("Error Cleared" & Space(35), 35)
    lstData.List(17) = lstData.List(17) & CStr(Value)
    
    Value = Cache.ErrorCorrectType
    lstData.AddItem Left("Error Correct Type" & Space(35), 35)
    Select Case Value
        Case 1: lstData.List(18) = lstData.List(18) & "Reserved"
        Case 2: lstData.List(18) = lstData.List(18) & "Other"
        Case 3: lstData.List(18) = lstData.List(18) & "Unknown"
        Case 4: lstData.List(18) = lstData.List(18) & "None"
        Case 5: lstData.List(18) = lstData.List(18) & "Parity"
        Case 6: lstData.List(18) = lstData.List(18) & "Single-bit ECC"
        Case 7: lstData.List(18) = lstData.List(18) & "Multi-bit ECC"
    End Select
    
    Value = Cache.ErrorData
    lstData.AddItem Left("Error Data" & Space(35), 35)
    lstData.List(19) = lstData.List(19) & CStr(Value)
    
    Value = Cache.ErrorDataOrder
    lstData.AddItem Left("Error Data Order" & Space(35), 35)
    Select Case Value
        Case 1: lstData.List(20) = lstData.List(20) & "Unknown"
        Case 1: lstData.List(20) = lstData.List(20) & "Least Significant Byte First"
        Case 1: lstData.List(20) = lstData.List(20) & "Most Significant Byte First"
    End Select
    
    Value = Cache.ErrorDescription
    lstData.AddItem Left("Error Description" & Space(35), 35)
    lstData.List(21) = lstData.List(21) & CStr(Value)
    
    Value = Cache.ErrorInfo
    lstData.AddItem Left("Error Info" & Space(35), 35)
    Select Case Value
        Case 1: lstData.List(22) = lstData.List(22) & "Other"
        Case 2: lstData.List(22) = lstData.List(22) & "Unknown"
        Case 3
            lstData.List(22) = lstData.List(22) & "OK"
            'If ok then clear the meaning of others
            'lstData.List(2) = Left(lstData.List(2), 35) 'AdditionalErrorData
            'lstData.List(15) = Left(lstData.List(15), 35) 'ErrorAccess
            'lstData.List(14) = Left(lstData.List(14), 35) 'ErrorAddres
            'lstData.List(24) = Left(lstData.List(24), 35) 'ErrorResolution
            'lstData.List(25) = Left(lstData.List(25), 35) 'ErrorTime
            'lstData.List(26) = Left(lstData.List(26), 35) 'ErrorTransferSize
        Case 4: lstData.List(22) = lstData.List(22) & "Bad Read"
        Case 5: lstData.List(22) = lstData.List(22) & "Parity Error"
        Case 6: lstData.List(22) = lstData.List(22) & "Single-Bit Error"
        Case 7: lstData.List(22) = lstData.List(22) & "Double-Bit Error"
        Case 8: lstData.List(22) = lstData.List(22) & "Multi-Bit Error"
        Case 9: lstData.List(22) = lstData.List(22) & "Nibble Error"
        Case 10: lstData.List(22) = lstData.List(22) & "Checksum Error"
        Case 11: lstData.List(22) = lstData.List(22) & "CRC Error"
        Case 12: lstData.List(22) = lstData.List(22) & "Undefined"
        Case 13: lstData.List(22) = lstData.List(22) & "Undefined"
        Case 14: lstData.List(22) = lstData.List(22) & "Undefined"
    End Select
    
    Value = Cache.ErrorMethodology
    lstData.AddItem Left("Error Methodology" & Space(35), 35)
    lstData.List(23) = lstData.List(23) & CStr(Value)
    
    Value = Cache.ErrorResolution
    lstData.AddItem Left("Error Resolution" & Space(35), 35)
    lstData.List(24) = lstData.List(24) & CStr(Value) & "bytes"
    
    Value = Cache.ErrorTime
    lstData.AddItem Left("Error Time" & Space(35), 35)
    lstData.List(25) = lstData.List(25) & CStr(Value)
    
    Value = Cache.ErrorTransferSize
    lstData.AddItem Left("Error Transfer Size" & Space(35), 35)
    lstData.List(26) = lstData.List(26) & CStr(Value) & "bits"
    'If CStr(Value) = "0" Then 'Clear others
    '    txtErrorData.Text = ""
    '    txtErrorDataOrder.Text = ""
    'End If

    Value = Cache.FlushTimer
    lstData.AddItem Left("Flush Timer" & Space(35), 35)
    lstData.List(27) = lstData.List(27) & CStr(Value) & "seconds"
    
    Value = Cache.InstallDate
    lstData.AddItem Left("Install Date" & Space(35), 35)
    lstData.List(28) = lstData.List(28) & CStr(Value)
    
    Value = Cache.InstalledSize
    lstData.AddItem Left("Installed Size" & Space(35), 35)
    lstData.List(29) = lstData.List(29) & CStr(Value) & "Kilobytes"
    
    Value = Cache.LastErrorCode
    lstData.AddItem Left("Last Error Code" & Space(35), 35)
    lstData.List(30) = lstData.List(30) & CStr(Value)
    
    Value = Cache.Level
    lstData.AddItem Left("Level" & Space(35), 35)
    Select Case Value
        Case 1: lstData.List(31) = lstData.List(31) & "Other"
        Case 2: lstData.List(31) = lstData.List(31) & "Unknown"
        Case 3: lstData.List(31) = lstData.List(31) & "Primary"
        Case 4: lstData.List(31) = lstData.List(31) & "Secondary"
        Case 5: lstData.List(31) = lstData.List(31) & "Tertiary"
    End Select
    
    Value = Cache.LineSize
    lstData.AddItem Left("Line Size" & Space(35), 35)
    lstData.List(32) = lstData.List(32) & CStr(Value) & "bytes"
    
    Value = Cache.Location
    lstData.AddItem Left("Location" & Space(35), 35)
    Select Case Value
        Case 1: lstData.List(33) = lstData.List(33) & "Internal"
        Case 2: lstData.List(33) = lstData.List(33) & "External"
        Case 3: lstData.List(33) = lstData.List(33) & "Reserved"
        Case 4: lstData.List(33) = lstData.List(33) & "Unknown"
    End Select
    
    Value = Cache.MaxCacheSize
    lstData.AddItem Left("Max Cache Size" & Space(35), 35)
    lstData.List(34) = lstData.List(34) & CStr(Value) & "Kilobytes"
    
    Value = Cache.Name
    lstData.AddItem Left("Name" & Space(35), 35)
    lstData.List(35) = lstData.List(35) & CStr(Value)
    
    Value = Cache.NumberOfBlocks
    lstData.AddItem Left("Number Of Blocks" & Space(35), 35)
    lstData.List(36) = lstData.List(36) & CStr(Value)
    
    Value = Cache.OtherErrorDescription
    lstData.AddItem Left("Other Error Description" & Space(35), 35)
    lstData.List(37) = lstData.List(37) & CStr(Value)
    
    Value = Cache.PNPDeviceID
    lstData.AddItem Left("PNP Device ID" & Space(35), 35)
    lstData.List(38) = lstData.List(38) & CStr(Value)
    
    Value = Cache.PowerManagementCapabilities
    For tmpInt = 0 To 7
        Select Case Value(tmpInt) 'Put in data
            Case 0: lstPowerManagementCapabilities.AddItem "Unknown"
            Case 1: lstPowerManagementCapabilities.AddItem "Not Supported"
            Case 2: lstPowerManagementCapabilities.AddItem "Disabled"
            Case 3: lstPowerManagementCapabilities.AddItem "Enabled"
            Case 4: lstPowerManagementCapabilities.AddItem "Power Saving Modes Entered Automatically"
            Case 5: lstPowerManagementCapabilities.AddItem "Power State Settable"
            Case 6: lstPowerManagementCapabilities.AddItem "Power Cycling Supported"
            Case 7: lstPowerManagementCapabilities.AddItem "Timed Power On Supported"
            Case Else: Exit For
        End Select
    Next tmpInt
    
    Value = Cache.PowerManagementSupported
    lstData.AddItem Left("Power Management Supported" & Space(35), 35)
    lstData.List(39) = lstData.List(39) & CStr(Value)
    
    Value = Cache.Purpose
    lstData.AddItem Left("Purpose" & Space(35), 35)
    lstData.List(40) = lstData.List(40) & CStr(Value)
    
    Value = Cache.ReadPolicy
    lstData.AddItem Left("Read Policy" & Space(35), 35)
    Select Case Value
        Case 1: lstData.List(41) = lstData.List(41) & "Other"
        Case 2: lstData.List(41) = lstData.List(41) & "Unknown"
        Case 3: lstData.List(41) = lstData.List(41) & "Read"
        Case 4: lstData.List(41) = lstData.List(41) & "Read-Ahead"
        Case 5: lstData.List(41) = lstData.List(41) & "Read and Read-Ahead"
    End Select
    
    Value = Cache.ReplacementPolicy
    lstData.AddItem Left("Replacement Policy" & Space(35), 35)
    Select Case Value
        Case 1: lstData.List(42) = lstData.List(42) & "Other"
        Case 2: lstData.List(42) = lstData.List(42) & "Unknown"
        Case 3: lstData.List(42) = lstData.List(42) & "Least Recently Used (LRU)"
        Case 4: lstData.List(42) = lstData.List(42) & "First In First Out (FIFO)"
        Case 5: lstData.List(42) = lstData.List(42) & "Last In First Out (LIFO)"
        Case 6: lstData.List(42) = lstData.List(42) & "Least Frequently Used (LFU)"
        Case 7: lstData.List(42) = lstData.List(42) & "Most Frequently Used (MFU)"
    End Select
    
    Value = Cache.StartingAddress
    lstData.AddItem Left("Starting Address" & Space(35), 35)
    lstData.List(43) = lstData.List(43) & CStr(Value) & "kilobytes"
    
    Value = Cache.Status
    lstData.AddItem Left("Status" & Space(35), 35)
    lstData.List(44) = lstData.List(44) & CStr(Value)
    
    Value = Cache.StatusInfo
    lstData.AddItem Left("Status Info" & Space(35), 35)
    Select Case Value
        Case 1: lstData.List(45) = lstData.List(45) & "Other"
        Case 2: lstData.List(45) = lstData.List(45) & "Unknown"
        Case 3: lstData.List(45) = lstData.List(45) & "Enabled"
        Case 4: lstData.List(45) = lstData.List(45) & "Disabled"
        Case 5: lstData.List(45) = lstData.List(45) & "Not Applicable"
    End Select
    
    Value = Cache.SupportedSRAM
    lstData.AddItem Left("Supported SRAM" & Space(35), 35)
    Select Case Value
        Case 1: lstData.List(46) = lstData.List(46) & "Other"
        Case 2: lstData.List(46) = lstData.List(46) & "Unknown"
        Case 3: lstData.List(46) = lstData.List(46) & "Non-Burst"
        Case 4: lstData.List(46) = lstData.List(46) & "Burst"
        Case 5: lstData.List(46) = lstData.List(46) & "Pipeline Burst"
        Case 6: lstData.List(46) = lstData.List(46) & "Synchronous"
        Case 7: lstData.List(46) = lstData.List(46) & "Asynchronous"
    End Select
    
    Value = Cache.SystemCreationClassName
    lstData.AddItem Left("System Creation Class Name" & Space(35), 35)
    lstData.List(47) = lstData.List(47) & CStr(Value)
    
    Value = Cache.SystemLevelAddress
    lstData.AddItem Left("System Level Address" & Space(35), 35)
    lstData.List(48) = lstData.List(48) & CStr(Value)
    
    Value = Cache.SystemName
    lstData.AddItem Left("System Name" & Space(35), 35)
    lstData.List(49) = lstData.List(49) & CStr(Value)
    
    Value = Cache.WritePolicy
    lstData.AddItem Left("Write Policy" & Space(35), 35)
    Select Case Value
        Case 1: lstData.List(50) = lstData.List(50) & "Other"
        Case 2: lstData.List(50) = lstData.List(50) & "Unknown"
        Case 3: lstData.List(50) = lstData.List(50) & "Write Back"
        Case 4: lstData.List(50) = lstData.List(50) & "Write Through"
        Case 5: lstData.List(50) = lstData.List(50) & "Varies with Address"
    End Select
    
    Me.MousePointer = vbNormal
End Sub
