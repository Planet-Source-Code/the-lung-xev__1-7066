VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Xev"
   ClientHeight    =   540
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   1560
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   540
   ScaleWidth      =   1560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuMain 
      Caption         =   ""
      Begin VB.Menu mnuHardware 
         Caption         =   "Hardware"
         Begin VB.Menu mnuCoolingDeviceC 
            Caption         =   "Cooling Device"
            Begin VB.Menu mnuFan 
               Caption         =   "Fan"
            End
         End
         Begin VB.Menu mnuInputDeviceC 
            Caption         =   "Input Device"
            Begin VB.Menu mnuKeyboard 
               Caption         =   "Keyboard"
            End
         End
         Begin VB.Menu mnuMassStorageC 
            Caption         =   "Mass Storage"
            Begin VB.Menu mnuCDROMDrive 
               Caption         =   "CDROM Drive"
            End
            Begin VB.Menu mnuFloppyDrive 
               Caption         =   "Floppy Drive"
            End
            Begin VB.Menu mnuLogicalDisk 
               Caption         =   "Logical Disk"
            End
         End
         Begin VB.Menu mnuMotherboardControllerPortC 
            Caption         =   "Motherboard/Controller/Port"
            Begin VB.Menu mnuBaseBoard 
               Caption         =   "Base Board"
            End
            Begin VB.Menu mnuBios 
               Caption         =   "Bios"
            End
            Begin VB.Menu mnuBus 
               Caption         =   "Bus"
            End
            Begin VB.Menu mnuCacheMemory 
               Caption         =   "Cache Memory"
            End
            Begin VB.Menu mnuParallelPort 
               Caption         =   "Parallel Port"
            End
            Begin VB.Menu mnuPhysicalMemory 
               Caption         =   "Physical Memory"
            End
            Begin VB.Menu mnuProcessor 
               Caption         =   "Processor"
            End
            Begin VB.Menu mnuSoundDevice 
               Caption         =   "Sound Device"
            End
            Begin VB.Menu mnuSystemMemoryResource 
               Caption         =   "System Memory Resource"
            End
         End
         Begin VB.Menu mnuNetworkingDeviceC 
            Caption         =   "Networking Device"
         End
         Begin VB.Menu mnuPowerC 
            Caption         =   "Power"
            Begin VB.Menu mnuVoltageProbe 
               Caption         =   "Voltage Probe"
            End
         End
         Begin VB.Menu mnuPrintingC 
            Caption         =   "Printing"
         End
         Begin VB.Menu mnuTelephonyC 
            Caption         =   "Telephony"
         End
         Begin VB.Menu mnuVideoMonitorC 
            Caption         =   "Video and Monitor"
            Begin VB.Menu mnuDesktopMonitor 
               Caption         =   "Desktop Monitor"
            End
            Begin VB.Menu mnuVideoConfiguration 
               Caption         =   "Video Configuration"
            End
            Begin VB.Menu mnuVideoController 
               Caption         =   "Video Controller"
            End
         End
      End
      Begin VB.Menu mnuOperatingSystem 
         Caption         =   "Operation System"
         Begin VB.Menu mnuCOMDCOMC 
            Caption         =   "COM/DCOM"
         End
         Begin VB.Menu mnuDesktopC 
            Caption         =   "Desktop"
         End
         Begin VB.Menu mnuDriversC 
            Caption         =   "Drivers"
         End
         Begin VB.Menu mnuFileSystemC 
            Caption         =   "File System"
         End
         Begin VB.Menu mnuMemoryPageFilesC 
            Caption         =   "Memory/Page Files"
         End
         Begin VB.Menu mnuMultimediaAudioVisualC 
            Caption         =   "Multimedia Audio/Visual"
         End
         Begin VB.Menu mnuNetworkingC 
            Caption         =   "Networking"
         End
         Begin VB.Menu mnuOperatingSystemSettingsC 
            Caption         =   "Operating System Settings"
            Begin VB.Menu mnuComputerSystem 
               Caption         =   "Computer System"
            End
         End
         Begin VB.Menu mnuProcessesC 
            Caption         =   "Processes"
            Begin VB.Menu mnuThread 
               Caption         =   "Thread"
            End
         End
         Begin VB.Menu mnuRegistryC 
            Caption         =   "Registry"
            Begin VB.Menu mnuRegistry 
               Caption         =   "Registry"
            End
         End
         Begin VB.Menu mnuSchedulerJobsC 
            Caption         =   "Scheduler Jobs"
         End
         Begin VB.Menu mnuSecurityC 
            Caption         =   "Security"
         End
         Begin VB.Menu mnuServicesC 
            Caption         =   "Services"
         End
         Begin VB.Menu mnuSharesC 
            Caption         =   "Shares"
         End
         Begin VB.Menu mnuStartMenuC 
            Caption         =   "Start Menu"
         End
         Begin VB.Menu mnuUsersC 
            Caption         =   "Users"
         End
         Begin VB.Menu mnuWindowsNTEventLogC 
            Caption         =   "Windows NT Event Log"
         End
      End
      Begin VB.Menu mnuSoftware 
         Caption         =   "Software"
      End
      Begin VB.Menu mnuBreak0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuOpenAll 
         Caption         =   "Open All"
      End
      Begin VB.Menu mnuBreak1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    On Error GoTo Errors
    
    Me.Hide
    
    If App.PrevInstance = True Then End 'If app run twice then exit
    
    Set Namespace = GetObject("winmgmts:") 'Login to root\cimv2

    'Add icon to system tray
    With NOTIFYICONDATA
        .cbSize = Len(NOTIFYICONDATA)
        .hwnd = frmMain.hwnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
        .uCallbackMessage = 512
        .hIcon = frmMain.Icon
        .szTip = frmMain.Caption + Chr(0) 'Tooltip text
    End With
    apiError = Shell_NotifyIcon(NIM_ADD, NOTIFYICONDATA)
    If apiError = 0 Then MsgBox "Failed" & vbCrLf & "Shell_NotifyIcon", vbInformation, "Error"
    
    Exit Sub
    
Errors:
    MsgBox "Critical error has occured.", vbExclamation, "Error" 'Warn
    End 'Exit here
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Subclass callback
    Dim tmpLong As Single
    tmpLong = X / Screen.TwipsPerPixelX
    
    Select Case tmpLong 'For system tray icon
        Case WM_LBUTTONUP
            apiError = SetForegroundWindow(Me.hwnd) 'Make sure its on top
            frmMain.PopupMenu mnuMain 'Popup menu
        Case WM_RBUTTONUP
            apiError = SetForegroundWindow(Me.hwnd) 'Make sure its on top
            frmMain.PopupMenu mnuMain 'Popup menu
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Program_End
End Sub

Private Sub mnuAbout_Click()
    frmMainAbout.Show
End Sub

Private Sub mnuBaseBoard_Click()
    frmBaseBoard.Show
End Sub

Private Sub mnuBios_Click()
    frmBios.Show
End Sub

Private Sub mnuBus_Click()
    frmBus.Show
End Sub

Private Sub mnuCacheMemory_Click()
    frmCacheMemory.Show
End Sub

Private Sub mnuCDROMDrive_Click()
    frmCDROMDrive.Show
End Sub

Private Sub mnuComputerSystem_Click()
    frmComputerSystem.Show
End Sub

Private Sub mnuDesktopMonitor_Click()
    frmDesktopMonitor.Show
End Sub

Private Sub mnuExit_Click()
    Program_End
End Sub

Private Sub mnuFan_Click()
    frmFan.Show
End Sub

Private Sub mnuFloppyDrive_Click()
    frmFloppyDrive.Show
End Sub

Private Sub mnuKeyboard_Click()
    frmKeyboard.Show
End Sub

Private Sub mnuLogicalDisk_Click()
    frmLogicalDisk.Show
End Sub

Private Sub mnuOpenAll_Click()
    'Show all manually
    frmBaseBoard.Show
    frmBios.Show
    frmBus.Show
    frmCacheMemory.Show
    frmCDROMDrive.Show
    frmComputerSystem.Show
    frmDesktopMonitor.Show
    frmFan.Show
    frmFloppyDrive.Show
    frmKeyboard.Show
    frmLogicalDisk.Show
    frmMainAbout.Show
    frmParallelPort.Show
    frmPhysicalMemory.Show
    frmProcessor.Show
    frmRegistry.Show
    frmSoundDevice.Show
    frmSystemMemoryResource.Show
    frmThread.Show
    frmVideoConfiguration.Show
    frmVideoController.Show
    frmVoltageProbe.Show
End Sub

Private Sub mnuParallelPort_Click()
    frmParallelPort.Show
End Sub

Private Sub mnuPhysicalMemory_Click()
    frmPhysicalMemory.Show
End Sub

Private Sub mnuProcessor_Click()
    frmProcessor.Show
End Sub

Private Sub mnuRegistry_Click()
    frmRegistry.Show
End Sub

Private Sub mnuSoundDevice_Click()
    frmSoundDevice.Show
End Sub

Private Sub mnuSystemMemoryResource_Click()
    frmSystemMemoryResource.Show
End Sub

Private Sub mnuThread_Click()
    frmThread.Show
End Sub

Private Sub mnuVideoConfiguration_Click()
    frmVideoConfiguration.Show
End Sub

Private Sub mnuVideoController_Click()
    frmVideoController.Show
End Sub

Private Sub mnuVoltageProbe_Click()
    frmVoltageProbe.Show
End Sub

'Its needed in 2 places so I created a propriatary function
Private Sub Program_End()
    'Remove icon from system tray
    With NOTIFYICONDATA
        .cbSize = Len(NOTIFYICONDATA)
        .hwnd = frmMain.hwnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
        .uCallbackMessage = vbNull
        .hIcon = frmMain.Icon
        .szTip = Chr(0) 'Clear
    End With
    apiError = Shell_NotifyIcon(NIM_DELETE, NOTIFYICONDATA)
    If apiError = 0 Then MsgBox "Failed" & vbCrLf & "Shell_NotifyIcon", vbInformation, "Error"
    
    End
End Sub
