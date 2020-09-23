VERSION 5.00
Begin VB.Form frmVideoConfiguration 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Video Configuration"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   Icon            =   "frmVideoConfiguration.frx":0000
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
   Begin VB.ListBox lstVideoConfiguration 
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
Attribute VB_Name = "frmVideoConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGetList_Click()
    On Error Resume Next
    
    Dim VideoConfiguration As SWbemObject
   
    'Clear current
    lstVideoConfiguration.Clear
    
    Me.MousePointer = vbHourglass
      
    'Enumerate the instances
    Dim VideoConfigurationSet As SWbemObjectSet
    Set VideoConfigurationSet = Namespace.InstancesOf("Win32_VideoConfiguration")
    
    For Each VideoConfiguration In VideoConfigurationSet
        ' Use the RelPath property of the instance path to display the disk
        lstVideoConfiguration.AddItem VideoConfiguration.Path_.RelPath
    Next

    Me.MousePointer = vbNormal
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub lstVideoConfiguration_Click()
    On Error Resume Next
    
    Dim SelectedItem As String
    Dim Value As Variant
    Dim VideoConfiguration As SWbemObject
    Dim tmpInt As Integer
    
    'Clear current
    lstData.Clear
    
    Me.MousePointer = vbHourglass
    
    SelectedItem = lstVideoConfiguration.List(lstVideoConfiguration.ListIndex)
    Set VideoConfiguration = Namespace.Get(SelectedItem)
    
    Value = VideoConfiguration.ActualColorResolution
    lstData.AddItem Left("Actual Color Resolution" & Space(35), 35)
    lstData.List(0) = lstData.List(0) & CStr(Value)

    Value = VideoConfiguration.AdapterChipType
    lstData.AddItem Left("Adapter Chip Type" & Space(35), 35)
    lstData.List(1) = lstData.List(1) & CStr(Value)

    Value = VideoConfiguration.AdapterCompatibility
    lstData.AddItem Left("Adapter Compatibility" & Space(35), 35)
    lstData.List(2) = lstData.List(2) & CStr(Value)

    Value = VideoConfiguration.AdapterDACType
    lstData.AddItem Left("Adapter DAC Type" & Space(35), 35)
    lstData.List(3) = lstData.List(3) & CStr(Value)
    
    Value = VideoConfiguration.AdapterDescription
    lstData.AddItem Left("Adapter Description" & Space(35), 35)
    lstData.List(4) = lstData.List(4) & CStr(Value)

    Value = VideoConfiguration.AdapterRAM
    lstData.AddItem Left("Adapter RAM" & Space(35), 35)
    lstData.List(5) = lstData.List(5) & CStr(Value) & "Bytes"

    Value = VideoConfiguration.AdapterType
    lstData.AddItem Left("Adapter Type" & Space(35), 35)
    lstData.List(6) = lstData.List(6) & CStr(Value)

    Value = VideoConfiguration.BitsPerPixel
    lstData.AddItem Left("Bits Per Pixel" & Space(35), 35)
    lstData.List(7) = lstData.List(7) & CStr(Value)

    Value = VideoConfiguration.Caption
    lstData.AddItem Left("Caption" & Space(35), 35)
    lstData.List(8) = lstData.List(8) & CStr(Value)

    Value = VideoConfiguration.ColorPlanes
    lstData.AddItem Left("Color Planes" & Space(35), 35)
    lstData.List(9) = lstData.List(9) & CStr(Value)

    Value = VideoConfiguration.ColorTableEntries
    lstData.AddItem Left("Color Table Entries" & Space(35), 35)
    lstData.List(10) = lstData.List(10) & CStr(Value)

    Value = VideoConfiguration.Description
    lstData.AddItem Left("Description" & Space(35), 35)
    lstData.List(11) = lstData.List(11) & CStr(Value)

    Value = VideoConfiguration.DeviceSpecificPens
    lstData.AddItem Left("Device Specific Pens" & Space(35), 35)
    lstData.List(12) = lstData.List(12) & CStr(Value)

    Value = VideoConfiguration.DriverDate
    lstData.AddItem Left("Driver Date" & Space(35), 35)
    lstData.List(13) = lstData.List(13) & CStr(Value)

    Value = VideoConfiguration.HorizontalResolution
    lstData.AddItem Left("Horizontal Resolution" & Space(35), 35)
    lstData.List(14) = lstData.List(14) & CStr(Value)

    Value = VideoConfiguration.InfFilename
    lstData.AddItem Left("Inf Filename" & Space(35), 35)
    lstData.List(15) = lstData.List(15) & CStr(Value)

    Value = VideoConfiguration.InfSection
    lstData.AddItem Left("Inf Section" & Space(35), 35)
    lstData.List(16) = lstData.List(16) & CStr(Value)

    Value = VideoConfiguration.InstalledDisplayDrivers
    lstData.AddItem Left("Installed Display Drivers" & Space(35), 35)
    lstData.List(17) = lstData.List(17) & CStr(Value)

    Value = VideoConfiguration.MonitorManufacturer
    lstData.AddItem Left("Monitor Manufacturer" & Space(35), 35)
    lstData.List(18) = lstData.List(18) & CStr(Value)

    Value = VideoConfiguration.MonitorType
    lstData.AddItem Left("Monitor Type" & Space(35), 35)
    lstData.List(19) = lstData.List(19) & CStr(Value)

    Value = VideoConfiguration.Name
    lstData.AddItem Left("Name" & Space(35), 35)
    lstData.List(20) = lstData.List(20) & CStr(Value)

    Value = VideoConfiguration.PixelsPerXLogicalInch
    lstData.AddItem Left("Pixels Per X Logical Inch" & Space(35), 35)
    lstData.List(21) = lstData.List(21) & CStr(Value)

    Value = VideoConfiguration.PixelsPerYLogicalInch
    lstData.AddItem Left("Pixels Per Y Logical Inch" & Space(35), 35)
    lstData.List(22) = lstData.List(22) & CStr(Value)

    Value = VideoConfiguration.RefreshRate
    lstData.AddItem Left("Refresh Rate" & Space(35), 35)
    lstData.List(23) = lstData.List(23) & CStr(Value)

    Value = VideoConfiguration.ScanMode
    lstData.AddItem Left("Scan Mode" & Space(35), 35)
    lstData.List(24) = lstData.List(24) & CStr(Value)

    Value = VideoConfiguration.ScreenHeight
    lstData.AddItem Left("Screen Height" & Space(35), 35)
    lstData.List(25) = lstData.List(25) & CStr(Value) & "Millimeters"

    Value = VideoConfiguration.ScreenWidth
    lstData.AddItem Left("Screen Width" & Space(35), 35)
    lstData.List(26) = lstData.List(26) & CStr(Value) & "Millimeters"

    Value = VideoConfiguration.SettingID
    lstData.AddItem Left("Setting ID" & Space(35), 35)
    lstData.List(27) = lstData.List(27) & CStr(Value)

    Value = VideoConfiguration.SystemPaletteEntries
    lstData.AddItem Left("System Palette Entries" & Space(35), 35)
    lstData.List(28) = lstData.List(28) & CStr(Value)

    Value = VideoConfiguration.VerticalResolution
    lstData.AddItem Left("Vertical Resolution" & Space(35), 35)
    lstData.List(29) = lstData.List(29) & CStr(Value)
    
    Me.MousePointer = vbNormal
End Sub
