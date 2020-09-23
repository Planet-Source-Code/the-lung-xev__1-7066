VERSION 5.00
Begin VB.Form frmVideoController 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Video Controller"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   Icon            =   "frmVideoController.frx":0000
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
   Begin VB.ListBox lstAcceleratorCapabilities 
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
   Begin VB.CommandButton cmdGetList 
      Caption         =   "Get List"
      Height          =   350
      Left            =   7800
      TabIndex        =   10
      Top             =   5040
      Width           =   975
   End
   Begin VB.ListBox lstVideo 
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
   Begin VB.Label lblPowerManagementCapabilities 
      Caption         =   "Power Management Capabilities"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   4680
      Width           =   2295
   End
   Begin VB.Label lblCapabilityDescriptions 
      Caption         =   "Accelerator Capabilities"
      Height          =   255
      Left            =   4560
      TabIndex        =   6
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Label lblAcceleratorCapabilities 
      Caption         =   "Accelerator Capabilities"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3840
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
   Begin VB.Label lblData 
      Caption         =   "Data"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   2055
   End
End
Attribute VB_Name = "frmVideoController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGetList_Click()
    On Error Resume Next
    
    Dim Video As SWbemObject
   
    'Clear current
    lstVideo.Clear
    
    Me.MousePointer = vbHourglass
      
    'Enumerate the instances
    Dim VideoSet As SWbemObjectSet
    Set VideoSet = Namespace.InstancesOf("Win32_VideoController")
    
    For Each Video In VideoSet
        ' Use the RelPath property of the instance path to display the disk
        lstVideo.AddItem Video.Path_.RelPath
    Next

    Me.MousePointer = vbNormal
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub lstVideo_Click()
    On Error Resume Next
    
    Dim SelectedItem As String
    Dim Value As Variant
    Dim Video As SWbemObject
    Dim tmpInt As Integer
    
    'Clear current
    lstData.Clear
    lstAcceleratorCapabilities.Clear
    lstCapabilityDescriptions.Clear
    lstPowerManagementCapabilities.Clear
    
    Me.MousePointer = vbHourglass
    
    SelectedItem = lstVideo.List(lstVideo.ListIndex)
    Set Video = Namespace.Get(SelectedItem)
    
    Value = Video.Availability
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

    Value = Video.AcceleratorCapabilities
    For tmpInt = 0 To 3
        Select Case Value(tmpInt) 'Put in data
            Case 0: lstAcceleratorCapabilities.AddItem "Unknown"
            Case 1: lstAcceleratorCapabilities.AddItem "Other"
            Case 2: lstAcceleratorCapabilities.AddItem "Graphics Accelerator"
            Case 3: lstAcceleratorCapabilities.AddItem "3D Accelerator"
            Case Else: Exit For
        End Select
    Next tmpInt

    Value = Video.AdapterCompatibility
    lstData.AddItem Left("Adapter Compatibility" & Space(35), 35)
    lstData.List(1) = lstData.List(1) & CStr(Value)
    
    Value = Video.AdapterDACType
    lstData.AddItem Left("Adapter DAC Type" & Space(35), 35)
    lstData.List(2) = lstData.List(2) & CStr(Value)

    Value = Video.AdapterRAM
    lstData.AddItem Left("Adapter RAM" & Space(35), 35)
    lstData.List(3) = lstData.List(3) & CStr(Value) & "Bytes"
    
    Value = Video.CapabilityDescriptions
    tmpInt = 0 'Reset
    Err.Number = 0 'Reset
    Do While Err.Number = 0 'Cycle through array
        lstCapabilityDescriptions.AddItem CStr(Value(tmpInt))
        tmpInt = tmpInt + 1 'Incremet
    Loop
    lstCapabilityDescriptions.RemoveItem tmpInt - 1 'Remove blank extra

    Value = Video.Caption
    lstData.AddItem Left("Caption" & Space(35), 35)
    lstData.List(4) = lstData.List(4) & CStr(Value)

    Value = Video.ColorTableEntries
    lstData.AddItem Left("ColorTableEntries" & Space(35), 35)
    lstData.List(5) = lstData.List(5) & CStr(Value)

    Value = Video.ConfigManagerErrorCode
    lstData.AddItem Left("Config Manager Error Code" & Space(35), 35)
    Select Case Value
        Case 0: lstData.List(6) = lstData.List(6) & "This device is working properly."
        Case 1: lstData.List(6) = lstData.List(6) & "This device is not configured correctly."
        Case 2: lstData.List(6) = lstData.List(6) & "Windows cannot load the driver for this device."
        Case 3: lstData.List(6) = lstData.List(6) & "The driver for this device might be corrupted, or your system may be running low on memory or other resources."
        Case 4: lstData.List(6) = lstData.List(6) & "This device is not working properly. One of its drivers or your registry might be corrupted."
        Case 5: lstData.List(6) = lstData.List(6) & "The driver for this device needs a resource that Windows cannot manage."
        Case 6: lstData.List(6) = lstData.List(6) & "The boot configuration for this device conflicts with other devices."
        Case 7: lstData.List(6) = lstData.List(6) & "Cannot filter."
        Case 8: lstData.List(6) = lstData.List(6) & "The driver loader for the device is missing."
        Case 9: lstData.List(6) = lstData.List(6) & "This device is not working properly because the controlling firmware is reporting the resources for the device incorrectly."
        Case 10: lstData.List(6) = lstData.List(6) & "This device cannot start."
        Case 11: lstData.List(6) = lstData.List(6) & "This device failed."
        Case 12: lstData.List(6) = lstData.List(6) & "This device cannot find enough free resources that it can use."
        Case 13: lstData.List(6) = lstData.List(6) & "Windows cannot verify this device's resources."
        Case 14: lstData.List(6) = lstData.List(6) & "This device cannot work properly until you restart your computer."
        Case 15: lstData.List(6) = lstData.List(6) & "This device is not working properly because there is probably a re-enumeration problem. "
        Case 16: lstData.List(6) = lstData.List(6) & "Windows cannot identify all the resources this device uses."
        Case 17: lstData.List(6) = lstData.List(6) & "This device is asking for an unknown resource type."
        Case 18: lstData.List(6) = lstData.List(6) & "Reinstall the drivers for this device."
        Case 19: lstData.List(6) = lstData.List(6) & "Your registry might be corrupted."
        Case 20: lstData.List(6) = lstData.List(6) & "System failure: Try changing the driver for this device. If that does not work, see your hardware documentation."
        Case 21: lstData.List(6) = lstData.List(6) & "Windows is removing this device."
        Case 22: lstData.List(6) = lstData.List(6) & "This device is disabled."
        Case 23: lstData.List(6) = lstData.List(6) & "System failure: Try changing the driver for this device. If that doesn't work, see your hardware documentation."
        Case 24: lstData.List(6) = lstData.List(6) & "This device is not present, is not working properly, or does not have all its drivers installed."
        Case 25: lstData.List(6) = lstData.List(6) & "Windows is still setting up this device."
        Case 26: lstData.List(6) = lstData.List(6) & "Windows is still setting up this device."
        Case 27: lstData.List(6) = lstData.List(6) & "This device does not have valid log configuration."
        Case 28: lstData.List(6) = lstData.List(6) & "The drivers for this device are not installed."
        Case 29: lstData.List(6) = lstData.List(6) & "This device is disabled because the firmware of the device did not give it the required resources."
        Case 30: lstData.List(6) = lstData.List(6) & "This device is using an Interrupt Request (IRQ) resource that another device is using."
        Case 31: lstData.List(6) = lstData.List(6) & "This device is not working properly because Windows cannot load the drivers required for this device."
    End Select

    Value = Video.CreationClassName
    lstData.AddItem Left("Creation Class Name" & Space(35), 35)
    lstData.List(7) = lstData.List(7) & CStr(Value)

    Value = Video.CurrentBitsPerPixel
    lstData.AddItem Left("Current Bits Per Pixel" & Space(35), 35)
    lstData.List(8) = lstData.List(8) & CStr(Value)

    Value = Video.CurrentHorizontalResolution
    lstData.AddItem Left("Current Horizontal Resolution" & Space(35), 35)
    lstData.List(9) = lstData.List(9) & CStr(Value) & "pixels"
    
    Value = Video.CurrentNumberOfColors
    lstData.AddItem Left("Current Number Of Colors" & Space(35), 35)
    lstData.List(10) = lstData.List(10) & CStr(Value)
    
    Value = Video.CurrentNumberOfColumns
    lstData.AddItem Left("Current Number Of Columns" & Space(35), 35)
    lstData.List(11) = lstData.List(11) & CStr(Value)

    Value = Video.CurrentNumberOfRows
    lstData.AddItem Left("Current Number Of Rows" & Space(35), 35)
    lstData.List(12) = lstData.List(12) & CStr(Value)

    Value = Video.CurrentRefreshRate
    lstData.AddItem Left("Current Refresh Rate" & Space(35), 35)
    lstData.List(13) = lstData.List(13) & CStr(Value) & "Hertz"
    
    Value = Video.CurrentScanMode
    lstData.AddItem Left("CurrentScanMode" & Space(35), 35)
    Select Case Value
        Case 1: lstData.List(14) = lstData.List(14) & "Other"
        Case 2: lstData.List(14) = lstData.List(14) & "Unknown"
        Case 3: lstData.List(14) = lstData.List(14) & "Interlaced"
        Case 4: lstData.List(14) = lstData.List(14) & "Non Interlaced"
    End Select

    Value = Video.CurrentVerticalResolution
    lstData.AddItem Left("Current Vertical Resolution" & Space(35), 35)
    lstData.List(15) = lstData.List(15) & CStr(Value) & "pixels"

    Value = Video.Description
    lstData.AddItem Left("Description" & Space(35), 35)
    lstData.List(16) = lstData.List(16) & CStr(Value)

    Value = Video.DeviceID
    lstData.AddItem Left("DeviceID" & Space(35), 35)
    lstData.List(17) = lstData.List(17) & CStr(Value)

    Value = Video.DeviceSpecificPens
    lstData.AddItem Left("Device Specific Pens" & Space(35), 35)
    lstData.List(18) = lstData.List(18) & CStr(Value)

    Value = Video.DitherType
    lstData.AddItem Left("DitherType" & Space(35), 35)
    Select Case Value
        Case 1: lstData.List(19) = lstData.List(19) & "No dithering"
        Case 2: lstData.List(19) = lstData.List(19) & "Dithering with a coarse brush"
        Case 3: lstData.List(19) = lstData.List(19) & "Dithering with a fine brush"
        Case 4: lstData.List(19) = lstData.List(19) & "Line art dithering"
        Case 5: lstData.List(19) = lstData.List(19) & "Device does gray scaling"
    End Select

    Value = Video.DriverDate
    lstData.AddItem Left("Driver Date" & Space(35), 35)
    lstData.List(20) = lstData.List(20) & CStr(Value)

    Value = Video.DriverVersion
    lstData.AddItem Left("Driver Version" & Space(35), 35)
    lstData.List(21) = lstData.List(21) & CStr(Value)

    Value = Video.ErrorCleared
    lstData.AddItem Left("Error Cleared" & Space(35), 35)
    lstData.List(22) = lstData.List(22) & CStr(Value)

    Value = Video.ErrorDescription
    lstData.AddItem Left("Error Description" & Space(35), 35)
    lstData.List(23) = lstData.List(23) & CStr(Value)

    Value = Video.ICMIntent
    lstData.AddItem Left("ICM Intent" & Space(35), 35)
    Select Case Value
        Case 1: lstData.List(24) = lstData.List(24) & "Saturation"
        Case 2: lstData.List(24) = lstData.List(24) & "Contrast"
        Case 3: lstData.List(24) = lstData.List(24) & "Exact Color"
    End Select

    Value = Video.ICMMethod
    lstData.AddItem Left("ICM Method" & Space(35), 35)
    Select Case Value
        Case 1: lstData.List(25) = lstData.List(25) & "Disabled"
        Case 2: lstData.List(25) = lstData.List(25) & "Windows"
        Case 3: lstData.List(25) = lstData.List(25) & "Device Driver"
        Case 3: lstData.List(25) = lstData.List(25) & "Destination Device"
    End Select

    Value = Video.InfFilename
    lstData.AddItem Left("Inf Filename" & Space(35), 35)
    lstData.List(26) = lstData.List(26) & CStr(Value)

    Value = Video.InfSection
    lstData.AddItem Left("Inf Section" & Space(35), 35)
    lstData.List(27) = lstData.List(27) & CStr(Value)
    
    Value = Video.InstallDate
    lstData.AddItem Left("Install Date" & Space(35), 35)
    lstData.List(28) = lstData.List(28) & CStr(Value)

    Value = Video.InstalledDisplayDrivers
    lstData.AddItem Left("Installed Display Drivers" & Space(35), 35)
    lstData.List(29) = lstData.List(29) & CStr(Value)

    Value = Video.LastErrorCode
    lstData.AddItem Left("Last Error Code" & Space(35), 35)
    lstData.List(30) = lstData.List(30) & CStr(Value)

    Value = Video.MaxMemorySupported
    lstData.AddItem Left("Max Memory Supported" & Space(35), 35)
    lstData.List(31) = lstData.List(31) & CStr(Value) & "bytes"

    Value = Video.MaxNumberControlled
    lstData.AddItem Left("Max Number Controlled" & Space(35), 35)
    lstData.List(32) = lstData.List(32) & CStr(Value)

    Value = Video.MaxRefreshRate
    lstData.AddItem Left("MaxRefreshRate" & Space(35), 35)
    lstData.List(33) = lstData.List(33) & CStr(Value) & "hertz"

    Value = Video.MinRefreshRate
    lstData.AddItem Left("Min Refresh Rate" & Space(35), 35)
    lstData.List(34) = lstData.List(34) & CStr(Value) & "hertz"

    Value = Video.Monochrome
    lstData.AddItem Left("Monochrome" & Space(35), 35)
    lstData.List(35) = lstData.List(35) & CStr(Value)

    Value = Video.Name
    lstData.AddItem Left("Name" & Space(35), 35)
    lstData.List(36) = lstData.List(36) & CStr(Value)

    Value = Video.NumberOfColorPlanes
    lstData.AddItem Left("Number Of Color Planes" & Space(35), 35)
    lstData.List(37) = lstData.List(37) & CStr(Value)

    Value = Video.NumberOfVideoPages
    lstData.AddItem Left("Number Of Video Pages" & Space(35), 35)
    lstData.List(38) = lstData.List(38) & CStr(Value)

    Value = Video.PNPDeviceID
    lstData.AddItem Left("PNP Device ID" & Space(35), 35)
    lstData.List(39) = lstData.List(39) & CStr(Value)

    Value = Video.PowerManagementCapabilities
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

    Value = Video.PowerManagementSupported
    lstData.AddItem Left("Power Management Supported" & Space(35), 35)
    lstData.List(40) = lstData.List(40) & CStr(Value)

    Value = Video.ProtocolSupported
    lstData.AddItem Left("Protocol Supported" & Space(35), 35)
    lstData.List(41) = lstData.List(41) & CStr(Value)

    Value = Video.ReservedSystemPaletteEntries
    lstData.AddItem Left("ReservedSystemPaletteEntries" & Space(35), 35)
    lstData.List(42) = lstData.List(42) & CStr(Value)

    Value = Video.SpecificationVersion
    lstData.AddItem Left("Specification Version" & Space(35), 35)
    lstData.List(43) = lstData.List(43) & CStr(Value)

    Value = Video.Status
    lstData.AddItem Left("Status" & Space(35), 35)
    lstData.List(44) = lstData.List(44) & CStr(Value)
    
    Value = Video.StatusInfo
    lstData.AddItem Left("Status Info" & Space(35), 35)
    Select Case Value
        Case 1: lstData.List(45) = lstData.List(45) & "Other"
        Case 2: lstData.List(45) = lstData.List(45) & "Unknown"
        Case 3: lstData.List(45) = lstData.List(45) & "Enabled"
        Case 4: lstData.List(45) = lstData.List(45) & "Disabled"
        Case 5: lstData.List(45) = lstData.List(45) & "Not Applicable"
    End Select

    Value = Video.SystemCreationClassName
    lstData.AddItem Left("System Creation Class Name" & Space(35), 35)
    lstData.List(46) = lstData.List(46) & CStr(Value)

    Value = Video.SystemName
    lstData.AddItem Left("System Name" & Space(35), 35)
    lstData.List(47) = lstData.List(47) & CStr(Value)

    Value = Video.SystemPaletteEntries
    lstData.AddItem Left("System Palette Entries" & Space(35), 35)
    lstData.List(48) = lstData.List(48) & CStr(Value)

    Value = Video.TimeOfLastReset
    lstData.AddItem Left("Time Of Last Reset" & Space(35), 35)
    lstData.List(49) = lstData.List(49) & CStr(Value)
    
    Value = Video.VideoArchitecture
    lstData.AddItem Left("Video Architecture" & Space(35), 35)
    Select Case Value
        Case 1: lstData.List(50) = lstData.List(50) & "Other"
        Case 2: lstData.List(50) = lstData.List(50) & "Unknown"
        Case 3: lstData.List(50) = lstData.List(50) & "CGA"
        Case 4: lstData.List(50) = lstData.List(50) & "EGA"
        Case 5: lstData.List(50) = lstData.List(50) & "VGA"
        Case 6: lstData.List(50) = lstData.List(50) & "SVGA"
        Case 7: lstData.List(50) = lstData.List(50) & "MDA"
        Case 8: lstData.List(50) = lstData.List(50) & "HGC"
        Case 9: lstData.List(50) = lstData.List(50) & "MCGA"
        Case 10: lstData.List(50) = lstData.List(50) & "8514A"
        Case 11: lstData.List(50) = lstData.List(50) & "XGA"
        Case 12: lstData.List(50) = lstData.List(50) & "Linear Frame Buffer"
        Case 160: lstData.List(50) = lstData.List(50) & "PC-98"
    End Select

    Value = Video.VideoMemoryType
    lstData.AddItem Left("Video Memory Type" & Space(35), 35)
    Select Case Value
        Case 0: lstData.List(51) = lstData.List(51) & "Other"
        Case 1: lstData.List(51) = lstData.List(51) & "Unknown"
        Case 2: lstData.List(51) = lstData.List(51) & "VRAM"
        Case 3: lstData.List(51) = lstData.List(51) & "DRAM"
        Case 4: lstData.List(51) = lstData.List(51) & "SRAM"
        Case 5: lstData.List(51) = lstData.List(51) & "WRAM"
        Case 6: lstData.List(51) = lstData.List(51) & "EDO RAM"
        Case 7: lstData.List(51) = lstData.List(51) & "Burst Synchronous DRAM"
        Case 8: lstData.List(51) = lstData.List(51) & "Pipelined Burst SRAM"
        Case 9: lstData.List(51) = lstData.List(51) & "CDRAM"
        Case 10: lstData.List(51) = lstData.List(51) & "3DRAM"
        Case 11: lstData.List(51) = lstData.List(51) & "SDRAM"
        Case 12: lstData.List(51) = lstData.List(51) & "SGRAM"
    End Select

    Value = Video.VideoMode
    lstData.AddItem Left("Video Mode" & Space(35), 35)
    lstData.List(52) = lstData.List(52) & CStr(Value)

    Value = Video.VideoModeDescription
    lstData.AddItem Left("Video Mode Description" & Space(35), 35)
    lstData.List(53) = lstData.List(53) & CStr(Value)

    Value = Video.VideoProcessor
    lstData.AddItem Left("Video Processor" & Space(35), 35)
    lstData.List(54) = lstData.List(54) & CStr(Value)
    
    Me.MousePointer = vbNormal
End Sub
