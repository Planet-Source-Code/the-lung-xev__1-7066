VERSION 5.00
Begin VB.Form frmCDROMDrive 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CDROM Drive"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   Icon            =   "frmCDROMDrive.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      Left            =   120
      TabIndex        =   9
      Top             =   4920
      Width           =   4215
   End
   Begin VB.CommandButton cmdGetList 
      Caption         =   "Get List"
      Height          =   350
      Left            =   7800
      TabIndex        =   10
      Top             =   5040
      Width           =   975
   End
   Begin VB.ListBox lstCDROMDrive 
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
   Begin VB.ListBox lstCapabilities 
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
   Begin VB.ListBox lstCapabilityDescriptions 
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
   Begin VB.Label lblPowerManagementCapabilities 
      Caption         =   "Power Management Capabilities"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   4680
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
   Begin VB.Label lblData 
      Caption         =   "Data"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label lblCapabilities 
      Caption         =   "Capabilities"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   2295
   End
   Begin VB.Label lblCapabilityDescriptions 
      Caption         =   "Capability Descriptions"
      Height          =   255
      Left            =   4560
      TabIndex        =   6
      Top             =   3840
      Width           =   2295
   End
End
Attribute VB_Name = "frmCDROMDrive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGetList_Click()
    On Error Resume Next
    
    Dim CDROMDrive As SWbemObject
   
    'Clear current
    lstCDROMDrive.Clear
    
    Me.MousePointer = vbHourglass
      
    'Enumerate the instances
    Dim CDROMDriveSet As SWbemObjectSet
    Set CDROMDriveSet = Namespace.InstancesOf("Win32_CDROMDrive")
    
    For Each CDROMDrive In CDROMDriveSet
        ' Use the RelPath property of the instance path to display the disk
        lstCDROMDrive.AddItem CDROMDrive.Path_.RelPath
    Next

    Me.MousePointer = vbNormal
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub lstCDROMDrive_Click()
    On Error Resume Next
    
    Dim SelectedItem As String
    Dim Value As Variant
    Dim CDROMDrive As SWbemObject
    Dim tmpInt As Integer
    
    'Clear current
    lstData.Clear
    lstCapabilities.Clear
    lstCapabilityDescriptions.Clear
    lstPowerManagementCapabilities.Clear

    Me.MousePointer = vbHourglass
    
    SelectedItem = lstCDROMDrive.List(lstCDROMDrive.ListIndex)
    Set CDROMDrive = Namespace.Get(SelectedItem)
    
    Value = CDROMDrive.Availability
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

    Value = CDROMDrive.Capabilities
    For tmpInt = 0 To 9
        Select Case Value(tmpInt) 'Put in data
            Case 0: lstCapabilities.AddItem "Unknown"
            Case 1: lstCapabilities.AddItem "Other"
            Case 2: lstCapabilities.AddItem "Sequential Access"
            Case 3: lstCapabilities.AddItem "Random Access"
            Case 4: lstCapabilities.AddItem "Supports Writing"
            Case 5: lstCapabilities.AddItem "Encryption"
            Case 6: lstCapabilities.AddItem "Compression"
            Case 7: lstCapabilities.AddItem "Supports Removable Media"
            Case 8: lstCapabilities.AddItem "Manual Cleaning"
            Case 9: lstCapabilities.AddItem "Automatic Cleaning"
            Case Else: Exit For
        End Select
    Next tmpInt
    
    Value = CDROMDrive.CapabilityDescriptions
    tmpInt = 0 'Reset
    Err.Number = 0 'Reset
    Do While Err.Number = 0 'Cycle through array
        lstCapabilityDescriptions.AddItem CStr(Value(tmpInt))
        tmpInt = tmpInt + 1 'Incremet
    Loop
    lstCapabilityDescriptions.RemoveItem tmpInt - 1 'Remove blank extra
    
    Value = CDROMDrive.Caption
    lstData.AddItem Left("Caption" & Space(35), 35)
    lstData.List(1) = lstData.List(1) & CStr(Value)

    Value = CDROMDrive.CompressionMethod
    lstData.AddItem Left("Compression Method" & Space(35), 35)
    lstData.List(2) = lstData.List(2) & CStr(Value)

    Value = CDROMDrive.ConfigManagerErrorCode
    lstData.AddItem Left("Config Manager Error Code" & Space(35), 35)
    Select Case Value
        Case 0: lstData.List(3) = lstData.List(3) & "This device is working properly."
        Case 1: lstData.List(3) = lstData.List(3) & "This device is not configured correctly."
        Case 2: lstData.List(3) = lstData.List(3) & "Windows cannot load the driver for this device."
        Case 3: lstData.List(3) = lstData.List(3) & "The driver for this device might be corrupted, or your system may be running low on memory or other resources."
        Case 4: lstData.List(3) = lstData.List(3) & "This device is not working properly. One of its drivers or your registry might be corrupted."
        Case 5: lstData.List(3) = lstData.List(3) & "The driver for this device needs a resource that Windows cannot manage."
        Case 6: lstData.List(3) = lstData.List(3) & "The boot configuration for this device conflicts with other devices."
        Case 7: lstData.List(3) = lstData.List(3) & "Cannot filter."
        Case 8: lstData.List(3) = lstData.List(3) & "The driver loader for the device is missing."
        Case 9: lstData.List(3) = lstData.List(3) & "This device is not working properly because the controlling firmware is reporting the resources for the device incorrectly."
        Case 10: lstData.List(3) = lstData.List(3) & "This device cannot start."
        Case 11: lstData.List(3) = lstData.List(3) & "This device failed."
        Case 12: lstData.List(3) = lstData.List(3) & "This device cannot find enough free resources that it can use."
        Case 13: lstData.List(3) = lstData.List(3) & "Windows cannot verify this device's resources."
        Case 14: lstData.List(3) = lstData.List(3) & "This device cannot work properly until you restart your computer."
        Case 15: lstData.List(3) = lstData.List(3) & "This device is not working properly because there is probably a re-enumeration problem. "
        Case 16: lstData.List(3) = lstData.List(3) & "Windows cannot identify all the resources this device uses."
        Case 17: lstData.List(3) = lstData.List(3) & "This device is asking for an unknown resource type."
        Case 18: lstData.List(3) = lstData.List(3) & "Reinstall the drivers for this device."
        Case 19: lstData.List(3) = lstData.List(3) & "Your registry might be corrupted."
        Case 20: lstData.List(3) = lstData.List(3) & "System failure: Try changing the driver for this device. If that does not work, see your hardware documentation."
        Case 21: lstData.List(3) = lstData.List(3) & "Windows is removing this device."
        Case 22: lstData.List(3) = lstData.List(3) & "This device is disabled."
        Case 23: lstData.List(3) = lstData.List(3) & "System failure: Try changing the driver for this device. If that doesn't work, see your hardware documentation."
        Case 24: lstData.List(3) = lstData.List(3) & "This device is not present, is not working properly, or does not have all its drivers installed."
        Case 25: lstData.List(3) = lstData.List(3) & "Windows is still setting up this device."
        Case 26: lstData.List(3) = lstData.List(3) & "Windows is still setting up this device."
        Case 27: lstData.List(3) = lstData.List(3) & "This device does not have valid log configuration."
        Case 28: lstData.List(3) = lstData.List(3) & "The drivers for this device are not installed."
        Case 29: lstData.List(3) = lstData.List(3) & "This device is disabled because the firmware of the device did not give it the required resources."
        Case 30: lstData.List(3) = lstData.List(3) & "This device is using an Interrupt Request (IRQ) resource that another device is using."
        Case 31: lstData.List(3) = lstData.List(3) & "This device is not working properly because Windows cannot load the drivers required for this device."
    End Select
    
    Value = CDROMDrive.ConfigManagerUserConfig
    lstData.AddItem Left("Config Manager User Config" & Space(35), 35)
    lstData.List(4) = lstData.List(4) & CStr(Value)
    
    Value = CDROMDrive.CreationClassName
    lstData.AddItem Left("Creation Class Name" & Space(35), 35)
    lstData.List(5) = lstData.List(5) & CStr(Value)

    Value = CDROMDrive.DefaultBlockSize
    lstData.AddItem Left("Default Block Size" & Space(35), 35)
    lstData.List(6) = lstData.List(6) & CStr(Value) & "bytes"

    Value = CDROMDrive.Description
    lstData.AddItem Left("Description" & Space(35), 35)
    lstData.List(7) = lstData.List(7) & CStr(Value)

    Value = CDROMDrive.DeviceID
    lstData.AddItem Left("Device ID" & Space(35), 35)
    lstData.List(8) = lstData.List(8) & CStr(Value)

    Value = CDROMDrive.Drive
    lstData.AddItem Left("Drive" & Space(35), 35)
    lstData.List(9) = lstData.List(9) & CStr(Value)

    Value = CDROMDrive.DriveIntegrity
    lstData.AddItem Left("Drive Integrity" & Space(35), 35)
    lstData.List(10) = lstData.List(10) & CStr(Value)

    Value = CDROMDrive.ErrorCleared
    lstData.AddItem Left("Error Cleared" & Space(35), 35)
    lstData.List(11) = lstData.List(11) & CStr(Value)

    Value = CDROMDrive.ErrorDescription
    lstData.AddItem Left("Error Description" & Space(35), 35)
    lstData.List(12) = lstData.List(12) & CStr(Value)

    Value = CDROMDrive.ErrorMethodology
    lstData.AddItem Left("Error Methodology" & Space(35), 35)
    lstData.List(13) = lstData.List(13) & CStr(Value)

    Value = CDROMDrive.FileSystemFlags
    lstData.AddItem Left("File System Flags" & Space(35), 35)
    lstData.List(14) = lstData.List(14) & CStr(Value)

    Value = CDROMDrive.FileSystemFlagsEx
    lstData.AddItem Left("File System Flags" & Space(35), 35)
    lstData.List(15) = lstData.List(15) & CStr(Value)

    Value = CDROMDrive.Id
    lstData.AddItem Left("Id" & Space(35), 35)
    lstData.List(16) = lstData.List(16) & CStr(Value)

    Value = CDROMDrive.InstallDate
    lstData.AddItem Left("Install Date" & Space(35), 35)
    lstData.List(17) = lstData.List(17) & CStr(Value)

    Value = CDROMDrive.LastErrorCode
    lstData.AddItem Left("Last Error Code" & Space(35), 35)
    lstData.List(18) = lstData.List(18) & CStr(Value)

    Value = CDROMDrive.Manufacturer
    lstData.AddItem Left("Manufacturer" & Space(35), 35)
    lstData.List(19) = lstData.List(19) & CStr(Value)

    Value = CDROMDrive.MaxBlockSize
    lstData.AddItem Left("Max Block Size" & Space(35), 35)
    lstData.List(20) = lstData.List(20) & CStr(Value) & "bytes"

    Value = CDROMDrive.MaximumComponentLength
    lstData.AddItem Left("Maximum Component Length" & Space(35), 35)
    lstData.List(21) = lstData.List(21) & CStr(Value)

    Value = CDROMDrive.MaxMediaSize
    lstData.AddItem Left("Max Media Size" & Space(35), 35)
    lstData.List(22) = lstData.List(22) & CStr(Value) & "kilobytes"

    Value = CDROMDrive.MediaLoaded
    lstData.AddItem Left("Media Loaded" & Space(35), 35)
    lstData.List(23) = lstData.List(23) & CStr(Value)

    Value = CDROMDrive.MediaType
    lstData.AddItem Left("Media Type" & Space(35), 35)
    lstData.List(24) = lstData.List(24) & CStr(Value)

    Value = CDROMDrive.MinBlockSize
    lstData.AddItem Left("Min Block Size" & Space(35), 35)
    lstData.List(25) = lstData.List(25) & CStr(Value) & "bytes"

    Value = CDROMDrive.Name
    lstData.AddItem Left("Name" & Space(35), 35)
    lstData.List(26) = lstData.List(26) & CStr(Value)

    Value = CDROMDrive.NeedsCleaning
    lstData.AddItem Left("Needs Cleaning" & Space(35), 35)
    lstData.List(27) = lstData.List(27) & CStr(Value)
    
    Value = CDROMDrive.NumberOfMediaSupported
    lstData.AddItem Left("Number Of Media Supported" & Space(35), 35)
    lstData.List(28) = lstData.List(28) & CStr(Value)

    Value = CDROMDrive.PNPDeviceID
    lstData.AddItem Left("PNP Device ID" & Space(35), 35)
    lstData.List(29) = lstData.List(29) & CStr(Value)

    Value = CDROMDrive.PowerManagementCapabilities
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

    Value = CDROMDrive.PowerManagementSupported
    lstData.AddItem Left("Power Management Supported" & Space(35), 35)
    lstData.List(30) = lstData.List(30) & CStr(Value)

    Value = CDROMDrive.RevisionLevel
    lstData.AddItem Left("Revision Level" & Space(35), 35)
    lstData.List(31) = lstData.List(31) & CStr(Value)
    
    Value = CDROMDrive.SCSIBus
    lstData.AddItem Left("SCSI Bus" & Space(35), 35)
    lstData.List(32) = lstData.List(32) & CStr(Value)

    Value = CDROMDrive.SCSILogicalUnit
    lstData.AddItem Left("SCSI Logical Unit" & Space(35), 35)
    lstData.List(33) = lstData.List(33) & CStr(Value)

    Value = CDROMDrive.SCSIPort
    lstData.AddItem Left("SCSI Port" & Space(35), 35)
    lstData.List(34) = lstData.List(34) & CStr(Value)

    Value = CDROMDrive.SCSITargetId
    lstData.AddItem Left("SCSI Target Id" & Space(35), 35)
    lstData.List(35) = lstData.List(35) & CStr(Value)

    Value = CDROMDrive.Size
    lstData.AddItem Left("Size" & Space(35), 35)
    lstData.List(36) = lstData.List(36) & CStr(Value) & "Bytes"

    Value = CDROMDrive.Status
    lstData.AddItem Left("Status" & Space(35), 35)
    lstData.List(37) = lstData.List(37) & CStr(Value)

    Value = CDROMDrive.StatusInfo
    lstData.AddItem Left("Status Info" & Space(35), 35)
    Select Case Value
        Case 1: lstData.List(38) = lstData.List(38) & "Other"
        Case 2: lstData.List(38) = lstData.List(38) & "Unknown"
        Case 3: lstData.List(38) = lstData.List(38) & "Enabled"
        Case 4: lstData.List(38) = lstData.List(38) & "Disabled"
        Case 5: lstData.List(38) = lstData.List(38) & "Not Applicable"
    End Select

    Value = CDROMDrive.SystemCreationClassName
    lstData.AddItem Left("System Creation Class Name" & Space(35), 35)
    lstData.List(39) = lstData.List(39) & CStr(Value)

    Value = CDROMDrive.SystemName
    lstData.AddItem Left("System Name" & Space(35), 35)
    lstData.List(40) = lstData.List(40) & CStr(Value)

    Value = CDROMDrive.TransferRate
    lstData.AddItem Left("Transfer Rate" & Space(35), 35)
    lstData.List(41) = lstData.List(41) & CStr(Value) & "KiloBytes per Second"

    Value = CDROMDrive.VolumeName
    lstData.AddItem Left("Volume Name" & Space(35), 35)
    lstData.List(42) = lstData.List(42) & CStr(Value)
    
    Value = CDROMDrive.VolumeSerialNumber
    lstData.AddItem Left("Volume Serial Number" & Space(35), 35)
    lstData.List(43) = lstData.List(43) & CStr(Value)
    
    Me.MousePointer = vbNormal
End Sub
