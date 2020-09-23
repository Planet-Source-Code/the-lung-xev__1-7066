VERSION 5.00
Begin VB.Form frmComputerSystem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Computer System"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   Icon            =   "frmComputerSystem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstSystemStartupOptions 
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
      TabIndex        =   13
      Top             =   5760
      Width           =   4215
   End
   Begin VB.ListBox lstSupportContactDescription 
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
      TabIndex        =   11
      Top             =   4920
      Width           =   4215
   End
   Begin VB.ListBox lstRoles 
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
   Begin VB.ListBox lstOEMStringArray 
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
   Begin VB.CommandButton cmdGetList 
      Caption         =   "Get List"
      Height          =   350
      Left            =   7800
      TabIndex        =   14
      Top             =   5880
      Width           =   975
   End
   Begin VB.ListBox lstSystem 
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
   Begin VB.Label lblSystemStartupOptions 
      Caption         =   "System Startup Options"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   5520
      Width           =   2055
   End
   Begin VB.Label lblSupportContactDescription 
      Caption         =   "Support Contact Description"
      Height          =   255
      Left            =   4560
      TabIndex        =   10
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Label lblRoles 
      Caption         =   "Roles"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   4680
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
   Begin VB.Label lblOEMStringArray 
      Caption         =   "OEM String Array"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3840
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
   Begin VB.Label lblList 
      Caption         =   "List"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmComputerSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGetList_Click()
    On Error Resume Next
    
    Dim System As SWbemObject
   
    'Clear current
    lstSystem.Clear
    
    Me.MousePointer = vbHourglass
      
    'Enumerate the instances
    Dim SystemSet As SWbemObjectSet
    Set SystemSet = Namespace.InstancesOf("Win32_ComputerSystem")
    
    For Each System In SystemSet
        ' Use the RelPath property of the instance path to display the disk
        lstSystem.AddItem System.Path_.RelPath
    Next

    Me.MousePointer = vbNormal
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub lstSystem_Click()
    On Error Resume Next
    
    Dim SelectedItem As String
    Dim Value As Variant
    Dim System As SWbemObject
    Dim tmpInt As Integer
    
    'Clear current
    lstData.Clear
    lstOEMStringArray.Clear
    lstPowerManagementCapabilities.Clear
    lstRoles.Clear
    lstSupportContactDescription.Clear
    lstSystemStartupOptions.Clear
    
    Me.MousePointer = vbHourglass
    
    SelectedItem = lstSystem.List(lstSystem.ListIndex)
    Set System = Namespace.Get(SelectedItem)
    
    Value = System.AdminPasswordStatus
    lstData.AddItem Left("Admin Password Status" & Space(35), 35)
    Select Case Value
        Case 1: lstData.List(0) = lstData.List(0) & "Disabled"
        Case 2: lstData.List(0) = lstData.List(0) & "Enabled"
        Case 3: lstData.List(0) = lstData.List(0) & "Not Implemented"
        Case 4: lstData.List(0) = lstData.List(0) & "Unknown"
    End Select

    Value = System.AutomaticResetBootOption
    lstData.AddItem Left("Automatic Reset Boot Option" & Space(35), 35)
    lstData.List(1) = lstData.List(1) & CStr(Value)

    Value = System.AutomaticResetCapability
    lstData.AddItem Left("Automatic Reset Capability" & Space(35), 35)
    lstData.List(2) = lstData.List(2) & CStr(Value)

    Value = System.BootOptionOnLimit
    lstData.AddItem Left("Boot Option On Limit" & Space(35), 35)
    Select Case Value
        Case 1: lstData.List(3) = lstData.List(3) & "Reserved"
        Case 2: lstData.List(3) = lstData.List(3) & "Operating system"
        Case 3: lstData.List(3) = lstData.List(3) & "System utilities"
        Case 4: lstData.List(3) = lstData.List(3) & "Do not reboot"
    End Select

    Value = System.BootOptionOnWatchDog
    lstData.AddItem Left("Boot Option On Watch Dog" & Space(35), 35)
    Select Case Value
        Case 1: lstData.List(4) = lstData.List(4) & "Reserved"
        Case 2: lstData.List(4) = lstData.List(4) & "Operating system"
        Case 3: lstData.List(4) = lstData.List(4) & "System utilities"
        Case 4: lstData.List(4) = lstData.List(4) & "Do not reboot"
    End Select

    Value = System.BootROMSupported
    lstData.AddItem Left("Boot ROM Supported" & Space(35), 35)
    lstData.List(5) = lstData.List(5) & CStr(Value)

    Value = System.BootupState
    lstData.AddItem Left("Bootup State" & Space(35), 35)
    lstData.List(6) = lstData.List(6) & CStr(Value)

    Value = System.Caption
    lstData.AddItem Left("Caption" & Space(35), 35)
    lstData.List(7) = lstData.List(7) & CStr(Value)

    Value = System.ChassisBootupState
    lstData.AddItem Left("Chassis Bootup State" & Space(35), 35)
    Select Case Value
        Case 1: lstData.List(8) = lstData.List(8) & "Other"
        Case 2: lstData.List(8) = lstData.List(8) & "Unknown"
        Case 3: lstData.List(8) = lstData.List(8) & "Safe"
        Case 4: lstData.List(8) = lstData.List(8) & "Warning"
        Case 5: lstData.List(8) = lstData.List(8) & "Critical"
        Case 6: lstData.List(8) = lstData.List(8) & "Non-recoverable"
    End Select

    Value = System.CreationClassName
    lstData.AddItem Left("Creation Class Name" & Space(35), 35)
    lstData.List(9) = lstData.List(9) & CStr(Value)

    Value = System.CurrentTimeZone
    lstData.AddItem Left("Current Time Zone" & Space(35), 35)
    lstData.List(10) = lstData.List(10) & CStr(Value)

    Value = System.DaylightInEffect
    lstData.AddItem Left("Daylight In Effect" & Space(35), 35)
    lstData.List(11) = lstData.List(11) & CStr(Value)

    Value = System.Description
    lstData.AddItem Left("Description" & Space(35), 35)
    lstData.List(12) = lstData.List(12) & CStr(Value)

    Value = System.Domain
    lstData.AddItem Left("Domain" & Space(35), 35)
    lstData.List(13) = lstData.List(13) & CStr(Value)

    Value = System.DomainRole
    lstData.AddItem Left("Domain Role" & Space(35), 35)
    Select Case Value
        Case 0: lstData.List(14) = lstData.List(14) & "Standalone Workstation"
        Case 1: lstData.List(14) = lstData.List(14) & "Member Workstation"
        Case 2: lstData.List(14) = lstData.List(14) & "Standalone Server"
        Case 3: lstData.List(14) = lstData.List(14) & "Member Server"
        Case 4: lstData.List(14) = lstData.List(14) & "Backup Domain Controller"
        Case 5: lstData.List(14) = lstData.List(14) & "Primary Domain Controller"
    End Select

    Value = System.FrontPanelResetStatus
    lstData.AddItem Left("Front Panel Reset Status" & Space(35), 35)
    Select Case Value
        Case 1: lstData.List(15) = lstData.List(15) & "Disabled"
        Case 2: lstData.List(15) = lstData.List(15) & "Enabled"
        Case 3: lstData.List(15) = lstData.List(15) & "Not Implemented"
        Case 4: lstData.List(15) = lstData.List(15) & "Unknown"
    End Select

    Value = System.InfraredSupported
    lstData.AddItem Left("Infrared Supported" & Space(35), 35)
    lstData.List(16) = lstData.List(16) & CStr(Value)

    Value = System.InitialLoadInfo
    lstData.AddItem Left("Initial Load Info" & Space(35), 35)
    lstData.List(17) = lstData.List(17) & CStr(Value)

    Value = System.InstallDate
    lstData.AddItem Left("Install Date" & Space(35), 35)
    lstData.List(18) = lstData.List(18) & CStr(Value)

    Value = System.KeyboardPasswordStatus
    lstData.AddItem Left("Keyboard Password Status" & Space(35), 35)
    Select Case Value
        Case 1: lstData.List(19) = lstData.List(19) & "Disabled"
        Case 2: lstData.List(19) = lstData.List(19) & "Enabled"
        Case 3: lstData.List(19) = lstData.List(19) & "Not Implemented"
        Case 4: lstData.List(19) = lstData.List(19) & "Unknown"
    End Select

    Value = System.LastLoadInfo
    lstData.AddItem Left("Last Load Info" & Space(35), 35)
    lstData.List(20) = lstData.List(20) & CStr(Value)

    Value = System.Manufacturer
    lstData.AddItem Left("Manufacturer" & Space(35), 35)
    lstData.List(21) = lstData.List(21) & CStr(Value)
    
    Value = System.Model
    lstData.AddItem Left("Model" & Space(35), 35)
    lstData.List(22) = lstData.List(22) & CStr(Value)

    Value = System.Name
    lstData.AddItem Left("Name" & Space(35), 35)
    lstData.List(23) = lstData.List(23) & CStr(Value)

    Value = System.NameFormat
    lstData.AddItem Left("Name Format" & Space(35), 35)
    lstData.List(24) = lstData.List(24) & CStr(Value)

    Value = System.NetworkServerModeEnabled
    lstData.AddItem Left("Network Server Mode Enabled" & Space(35), 35)
    lstData.List(25) = lstData.List(25) & CStr(Value)

    Value = System.NumberOfProcessors
    lstData.AddItem Left("Number Of Processors" & Space(35), 35)
    lstData.List(26) = lstData.List(26) & CStr(Value)

    Value = System.OEMLogoBitmap
    lstData.AddItem Left("OEM Logo Bitmap" & Space(35), 35)
    lstData.List(27) = lstData.List(27) & CStr(Value)

    Value = System.OEMStringArray
    tmpInt = 0 'Reset
    Err.Number = 0 'Reset
    Do While Err.Number = 0 'Cycle through array
        lstOEMStringArray.AddItem CStr(Value(tmpInt))
        tmpInt = tmpInt + 1 'Incremet
    Loop
    lstOEMStringArray.RemoveItem tmpInt - 1 'Remove blank extra

    Value = System.PauseAfterReset
    lstData.AddItem Left("Pause After Reset" & Space(35), 35)
    lstData.List(28) = lstData.List(28) & CStr(Value) & "Milliseconds"

    Value = System.PowerManagementCapabilities
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
    
    Value = System.PowerManagementSupported
    lstData.AddItem Left("Power Management Supported" & Space(35), 35)
    lstData.List(29) = lstData.List(29) & CStr(Value)

    Value = System.PowerOnPasswordStatus
    lstData.AddItem Left("Power On Password Status" & Space(35), 35)
    Select Case Value
        Case 1: lstData.List(30) = lstData.List(30) & "Disabled"
        Case 2: lstData.List(30) = lstData.List(30) & "Enabled"
        Case 3: lstData.List(30) = lstData.List(30) & "Not Implemented"
        Case 4: lstData.List(30) = lstData.List(30) & "Unknown"
    End Select

    Value = System.PowerState
    lstData.AddItem Left("Power State" & Space(35), 35)
    lstData.List(31) = lstData.List(31) & CStr(Value)

    Value = System.PowerSupplyState
    lstData.AddItem Left("Power Supply State" & Space(35), 35)
    Select Case Value
        Case 1: lstData.List(32) = lstData.List(32) & "Other"
        Case 2: lstData.List(32) = lstData.List(32) & "Unknown"
        Case 3: lstData.List(32) = lstData.List(32) & "Safe"
        Case 4: lstData.List(32) = lstData.List(32) & "Warning"
        Case 4: lstData.List(32) = lstData.List(32) & "Critical"
        Case 4: lstData.List(32) = lstData.List(32) & "Non-recoverable"
    End Select

    Value = System.PrimaryOwnerContact
    lstData.AddItem Left("Primary Owner Contact" & Space(35), 35)
    lstData.List(33) = lstData.List(33) & CStr(Value)

    Value = System.PrimaryOwnerName
    lstData.AddItem Left("Primary Owner Name" & Space(35), 35)
    lstData.List(34) = lstData.List(34) & CStr(Value)

    Value = System.ResetCapability
    lstData.AddItem Left("Reset Capability" & Space(35), 35)
    Select Case Value
        Case 1: lstData.List(35) = lstData.List(35) & "Other"
        Case 2: lstData.List(35) = lstData.List(35) & "Unknown"
        Case 3: lstData.List(35) = lstData.List(35) & "Disabled"
        Case 4: lstData.List(35) = lstData.List(35) & "Enabled"
        Case 5: lstData.List(35) = lstData.List(35) & "Not Implemented"
    End Select

    Value = System.ResetCount
    lstData.AddItem Left("Reset Count" & Space(35), 35)
    lstData.List(36) = lstData.List(36) & CStr(Value)
    
    Value = System.ResetLimit
    lstData.AddItem Left("Reset Limit" & Space(35), 35)
    lstData.List(37) = lstData.List(37) & CStr(Value)

    Value = System.Roles
    tmpInt = 0 'Reset
    Err.Number = 0 'Reset
    Do While Err.Number = 0 'Cycle through array
        lstRoles.AddItem CStr(Value(tmpInt))
        tmpInt = tmpInt + 1 'Incremet
    Loop
    lstRoles.RemoveItem tmpInt - 1 'Remove blank extra

    Value = System.Status
    lstData.AddItem Left("Status" & Space(35), 35)
    lstData.List(38) = lstData.List(38) & CStr(Value)
    
    Value = System.Status
    lstData.AddItem Left("Status" & Space(35), 35)
    lstData.List(39) = lstData.List(39) & CStr(Value)
    
    Value = System.SupportContactDescription
    tmpInt = 0 'Reset
    Err.Number = 0 'Reset
    Do While Err.Number = 0 'Cycle through array
        lstSupportContactDescription.AddItem CStr(Value(tmpInt))
        tmpInt = tmpInt + 1 'Incremet
    Loop
    lstSupportContactDescription.RemoveItem tmpInt - 1 'Remove blank extra

    Value = System.SystemStartupDelay
    lstData.AddItem Left("System Startup Delay" & Space(35), 35)
    lstData.List(40) = lstData.List(40) & CStr(Value) & "Seconds"

    Value = System.SystemStartupOptions
    tmpInt = 0 'Reset
    Err.Number = 0 'Reset
    Do While Err.Number = 0 'Cycle through array
        lstSystemStartupOptions.AddItem CStr(Value(tmpInt))
        tmpInt = tmpInt + 1 'Incremet
    Loop
    lstSystemStartupOptions.RemoveItem tmpInt - 1 'Remove blank extra

    Value = System.SystemStartupSetting
    lstData.AddItem Left("System Startup Setting" & Space(35), 35)
    lstData.List(41) = lstData.List(41) & CStr(Value)

    Value = System.SystemType
    lstData.AddItem Left("System Type" & Space(35), 35)
    lstData.List(42) = lstData.List(42) & CStr(Value)

    Value = System.ThermalState
    lstData.AddItem Left("Thermal State" & Space(35), 35)
    Select Case Value
        Case 1: lstData.List(43) = lstData.List(43) & "Other"
        Case 2: lstData.List(43) = lstData.List(43) & "Unknown"
        Case 3: lstData.List(43) = lstData.List(43) & "Safe"
        Case 4: lstData.List(43) = lstData.List(43) & "Warning"
        Case 5: lstData.List(43) = lstData.List(43) & "Critical"
        Case 6: lstData.List(43) = lstData.List(43) & "Non-recoverable"
    End Select

    Value = System.TotalPhysicalMemory
    lstData.AddItem Left("Total Physical Memory" & Space(35), 35)
    lstData.List(44) = lstData.List(44) & CStr(Value) & "Bytes"

    Value = System.UserName
    lstData.AddItem Left("User Name" & Space(35), 35)
    lstData.List(45) = lstData.List(45) & CStr(Value)

    Value = System.WakeUpType
    lstData.AddItem Left("Wake Up Type" & Space(35), 35)
    Select Case Value
        Case 1: lstData.List(46) = lstData.List(46) & "Reserved"
        Case 2: lstData.List(46) = lstData.List(46) & "Other"
        Case 3: lstData.List(46) = lstData.List(46) & "Unknown"
        Case 4: lstData.List(46) = lstData.List(46) & "APM Timer"
        Case 5: lstData.List(46) = lstData.List(46) & "Modem Ring"
        Case 6: lstData.List(46) = lstData.List(46) & "LAN Remote"
        Case 7: lstData.List(46) = lstData.List(46) & "Power Switch"
        Case 8: lstData.List(46) = lstData.List(46) & "PCI PME#"
        Case 9: lstData.List(46) = lstData.List(46) & "AC Power Restored"
    End Select
    
    Me.MousePointer = vbNormal
End Sub
