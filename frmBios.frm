VERSION 5.00
Begin VB.Form frmBios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bios"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   Icon            =   "frmBios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
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
   Begin VB.ListBox lstListOfLanguages 
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
      TabIndex        =   7
      Top             =   5520
      Width           =   4215
   End
   Begin VB.ListBox lstCharacteristics 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   120
      TabIndex        =   5
      Top             =   4080
      Width           =   8655
   End
   Begin VB.ListBox lstBios 
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
      Top             =   5640
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
   Begin VB.Label lblListOfLanguages 
      Caption         =   "List Of Languages"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   5280
      Width           =   2055
   End
   Begin VB.Label lblCharacteristics 
      Caption         =   "Characteristics"
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
End
Attribute VB_Name = "frmBios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGetList_Click()
    On Error Resume Next
    
    Dim Bios As SWbemObject
   
    'Clear current
    lstBios.Clear
    
    Me.MousePointer = vbHourglass
      
    'Enumerate the instances
    Dim BiosSet As SWbemObjectSet
    Set BiosSet = Namespace.InstancesOf("Win32_BIOS")
    
    For Each Bios In BiosSet
        ' Use the RelPath property of the instance path to display the disk
        lstBios.AddItem Bios.Path_.RelPath
    Next

    Me.MousePointer = vbNormal
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub lstBios_Click()
    On Error Resume Next
    
    Dim SelectedItem As String
    Dim Value As Variant
    Dim Bios As SWbemObject
    Dim tmpInt As Integer

    'Clear current
    lstData.Clear
    lstCharacteristics.Clear
    lstListOfLanguages.Clear
    
    Me.MousePointer = vbHourglass

    SelectedItem = lstBios.List(lstBios.ListIndex)
    Set Bios = Namespace.Get(SelectedItem)
    
    Value = Bios.BiosCharacteristics
    For tmpInt = 0 To 49
        Select Case Value(tmpInt) 'Put in data
            Case 0: lstCharacteristics.AddItem "Reserved"
            Case 1: lstCharacteristics.AddItem "Reserved"
            Case 2: lstCharacteristics.AddItem "Unknown"
            Case 3: lstCharacteristics.AddItem "BIOS Characteristics Not Supported"
            Case 4: lstCharacteristics.AddItem "ISA is supported"
            Case 5: lstCharacteristics.AddItem "MCA is supported"
            Case 6: lstCharacteristics.AddItem "EISA is supported"
            Case 7: lstCharacteristics.AddItem "PCI is supported"
            Case 8: lstCharacteristics.AddItem "PC Card (PCMCIA) is supported"
            Case 9: lstCharacteristics.AddItem "Plug and Play is supported"
            Case 10: lstCharacteristics.AddItem "APM is supported"
            Case 11: lstCharacteristics.AddItem "BIOS is Upgradeable (Flash)"
            Case 12: lstCharacteristics.AddItem "BIOS shadowing is allowed"
            Case 13: lstCharacteristics.AddItem "VL-VESA is supported"
            Case 14: lstCharacteristics.AddItem "ESCD support is available"
            Case 15: lstCharacteristics.AddItem "Boot from CD is supported"
            Case 16: lstCharacteristics.AddItem "Selectable Boot is supported"
            Case 17: lstCharacteristics.AddItem "BIOS ROM is socketed"
            Case 18: lstCharacteristics.AddItem "Boot From PC Card (PCMCIA) is supported"
            Case 19: lstCharacteristics.AddItem "EDD (Enhanced Disk Drive) Specification is supported"
            Case 20: lstCharacteristics.AddItem "Int 13h - Japanese Floppy for NEC 9800 1.2mb (3.5, 1k Bytes/Sector, 360 RPM) is supported"
            Case 21: lstCharacteristics.AddItem "Int 13h - Japanese Floppy for Toshiba 1.2mb (3.5, 360 RPM) is supported"
            Case 22: lstCharacteristics.AddItem "Int 13h - 5.25 / 360 KB Floppy Services are supported"
            Case 23: lstCharacteristics.AddItem "Int 13h - 5.25 /1.2MB Floppy Services are supported"
            Case 24: lstCharacteristics.AddItem "Int 13h - 3.5 / 720 KB Floppy Services are supported"
            Case 25: lstCharacteristics.AddItem "Int 13h - 3.5 / 2.88 MB Floppy Services are supported"
            Case 26: lstCharacteristics.AddItem "Int 5h, Print Screen Service is supported"
            Case 27: lstCharacteristics.AddItem "Int 9h, 8042 Keyboard services are supported"
            Case 28: lstCharacteristics.AddItem "Int 14h, Serial Services are supported"
            Case 29: lstCharacteristics.AddItem "Int 17h, printer services are supported"
            Case 30: lstCharacteristics.AddItem "Int 10h, CGA/Mono Video Services are supported"
            Case 31: lstCharacteristics.AddItem "NEC PC-98"
            Case 32: lstCharacteristics.AddItem "ACPI supported"
            Case 33: lstCharacteristics.AddItem "USB Legacy is supported"
            Case 34: lstCharacteristics.AddItem "AGP is supported"
            Case 35: lstCharacteristics.AddItem "I2O boot is supported"
            Case 36: lstCharacteristics.AddItem "LS-120 boot is supported"
            Case 37: lstCharacteristics.AddItem "ATAPI ZIP Drive boot is supported"
            Case 38: lstCharacteristics.AddItem "1394 boot is supported"
            Case 39: lstCharacteristics.AddItem "Smart Battery supported"
            Case Else: Exit For
        End Select
    Next tmpInt
    
    Value = Bios.BuildNumber
    lstData.AddItem Left("Build Number" & Space(35), 35)
    lstData.List(0) = lstData.List(0) & CStr(Value)
    
    Value = Bios.Caption
    lstData.AddItem Left("Caption" & Space(35), 35)
    lstData.List(1) = lstData.List(1) & CStr(Value)
    
    Value = Bios.CodeSet
    lstData.AddItem Left("Code Set" & Space(35), 35)
    lstData.List(2) = lstData.List(2) & CStr(Value)
    
    Value = Bios.CurrentLanguage
    lstData.AddItem Left("Current Language" & Space(35), 35)
    lstData.List(3) = lstData.List(3) & CStr(Value)
    
    Value = Bios.Description
    lstData.AddItem Left("Description" & Space(35), 35)
    lstData.List(4) = lstData.List(4) & CStr(Value)
    
    Value = Bios.IdentificationCode
    lstData.AddItem Left("Identification Code" & Space(35), 35)
    lstData.List(5) = lstData.List(5) & CStr(Value)
    
    Value = Bios.InstallableLanguages
    lstData.AddItem Left("Installable Languages" & Space(35), 35)
    lstData.List(6) = lstData.List(6) & CStr(Value)
    
    Value = Bios.InstallDate
    lstData.AddItem Left("Install Date" & Space(35), 35)
    lstData.List(7) = lstData.List(7) & CStr(Value)
    
    Value = Bios.LanguageEdition
    lstData.AddItem Left("Language Edition" & Space(35), 35)
    lstData.List(8) = lstData.List(8) & CStr(Value)
    
    Value = Bios.ListOfLanguages
    tmpInt = 0 'Reset
    Err.Number = 0 'Reset
    Do While Err.Number = 0 'Cycle through array
        lstListOfLanguages.AddItem CStr(Value(tmpInt))
        tmpInt = tmpInt + 1 'Incremet
    Loop
    lstListOfLanguages.RemoveItem tmpInt - 1 'Remove blank extra
    
    Value = Bios.Manufacturer
    lstData.AddItem Left("Manufacturer" & Space(35), 35)
    lstData.List(9) = lstData.List(9) & CStr(Value)
    
    Value = Bios.Name
    lstData.AddItem Left("Name" & Space(35), 35)
    lstData.List(10) = lstData.List(10) & CStr(Value)
    
    Value = Bios.OtherTargetOS
    lstData.AddItem Left("Other Target OS" & Space(35), 35)
    lstData.List(11) = lstData.List(11) & CStr(Value)
    
    Value = Bios.PrimaryBIOS
    lstData.AddItem Left("Primary BIOS" & Space(35), 35)
    lstData.List(12) = lstData.List(12) & CStr(Value)
    
    Value = Bios.ReleaseDate
    lstData.AddItem Left("Release Date" & Space(35), 35)
    lstData.List(13) = lstData.List(13) & CStr(Value)
    
    Value = Bios.SerialNumber
    lstData.AddItem Left("Serial Number" & Space(35), 35)
    lstData.List(14) = lstData.List(14) & CStr(Value)
    
    Value = Bios.SMBIOSBIOSVersion
    lstData.AddItem Left("SMBIOS BIOS Version" & Space(35), 35)
    lstData.List(15) = lstData.List(15) & CStr(Value)
    
    Value = Bios.SMBIOSMajorVersion
    lstData.AddItem Left("SMBIOS Major Version" & Space(35), 35)
    lstData.List(16) = lstData.List(16) & CStr(Value)
    
    Value = Bios.SMBIOSMinorVersion
    lstData.AddItem Left("SMBIOS Minor Version" & Space(35), 35)
    lstData.List(17) = lstData.List(17) & CStr(Value)
    
    Value = Bios.SMBIOSPresent
    lstData.AddItem Left("SMBIOS Present" & Space(35), 35)
    lstData.List(18) = lstData.List(18) & CStr(Value)
    
    Value = Bios.SoftwareElementID
    lstData.AddItem Left("Software Element ID" & Space(35), 35)
    lstData.List(19) = lstData.List(19) & CStr(Value)
    
    Value = Bios.SoftwareElementState
    lstData.AddItem Left("Software Element State" & Space(35), 35)
    Select Case Value
        Case 1: lstData.List(20) = lstData.List(20) & "Deployable"
        Case 2: lstData.List(20) = lstData.List(20) & "Installable"
        Case 3: lstData.List(20) = lstData.List(20) & "Executable"
        Case 4: lstData.List(20) = lstData.List(20) & "Running"
    End Select
    
    Value = Bios.Status
    lstData.AddItem Left("Status" & Space(35), 35)
    lstData.List(21) = lstData.List(21) & CStr(Value)
    
    Value = Bios.TargetOperatingSystem
    lstData.AddItem Left("Target Operating System" & Space(35), 35)
    lstData.List(22) = lstData.List(22) & CStr(Value)
    
    Value = Bios.Version
    lstData.AddItem Left("Version" & Space(35), 35)
    lstData.List(23) = lstData.List(23) & CStr(Value)
    
    Me.MousePointer = vbNormal
End Sub
