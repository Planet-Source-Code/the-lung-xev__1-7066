VERSION 5.00
Begin VB.Form frmBaseBoard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Base Board"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   Icon            =   "frmBaseBoard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstConfigOptions 
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
      TabIndex        =   6
      Top             =   4200
      Width           =   975
   End
   Begin VB.ListBox lstBaseBoard 
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
   Begin VB.Label lblConfigOptions 
      Caption         =   "Config Options"
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
Attribute VB_Name = "frmBaseBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGetList_Click()
    On Error Resume Next
    
    Dim BaseBoard As SWbemObject
   
    'Clear current
    lstBaseBoard.Clear
    
    Me.MousePointer = vbHourglass
      
    'Enumerate the instances
    Dim BaseBoardSet As SWbemObjectSet
    Set BaseBoardSet = Namespace.InstancesOf("Win32_BaseBoard")
    
    For Each BaseBoard In BaseBoardSet
        ' Use the RelPath property of the instance path to display the disk
        lstBaseBoard.AddItem BaseBoard.Path_.RelPath
    Next

    Me.MousePointer = vbNormal
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub lstBaseBoard_Click()
    On Error Resume Next
    
    Dim SelectedItem As String
    Dim Value As Variant
    Dim BaseBoard As SWbemObject
    Dim tmpInt As Integer
    
    'Clear current
    lstData.Clear
    lstConfigOptions.Clear
    
    Me.MousePointer = vbHourglass
    
    SelectedItem = lstBaseBoard.List(lstBaseBoard.ListIndex)
    Set BaseBoard = Namespace.Get(SelectedItem)
    
    Value = BaseBoard.Caption
    lstData.AddItem Left("Caption" & Space(35), 35)
    lstData.List(0) = lstData.List(0) & CStr(Value)

    Value = BaseBoard.ConfigOptions
    tmpInt = 0 'Reset
    Err.Number = 0 'Reset
    Do While Err.Number = 0 'Cycle through array
        lstConfigOptions.AddItem CStr(Value(tmpInt))
        tmpInt = tmpInt + 1 'Incremet
    Loop
    lstConfigOptions.RemoveItem tmpInt - 1 'Remove blank extra

    Value = BaseBoard.CreationClassName
    lstData.AddItem Left("Creation Class Name" & Space(35), 35)
    lstData.List(1) = lstData.List(1) & CStr(Value)

    Value = BaseBoard.Depth
    lstData.AddItem Left("Depth" & Space(35), 35)
    lstData.List(2) = lstData.List(2) & CStr(Value) & "inches"

    Value = BaseBoard.Description
    lstData.AddItem Left("Description" & Space(35), 35)
    lstData.List(3) = lstData.List(3) & CStr(Value)

    Value = BaseBoard.Height
    lstData.AddItem Left("Height" & Space(35), 35)
    lstData.List(4) = lstData.List(4) & CStr(Value) & "inches"

    Value = BaseBoard.HostingBoard
    lstData.AddItem Left("Hosting Board" & Space(35), 35)
    lstData.List(5) = lstData.List(5) & CStr(Value)
    
    Value = BaseBoard.HotSwappable
    lstData.AddItem Left("Hot Swappable" & Space(35), 35)
    lstData.List(6) = lstData.List(6) & CStr(Value)

    Value = BaseBoard.InstallDate
    lstData.AddItem Left("Install Date" & Space(35), 35)
    lstData.List(7) = lstData.List(7) & CStr(Value)
    
    Value = BaseBoard.Manufacturer
    lstData.AddItem Left("Manufacturer" & Space(35), 35)
    lstData.List(8) = lstData.List(8) & CStr(Value)

    Value = BaseBoard.Model
    lstData.AddItem Left("Model" & Space(35), 35)
    lstData.List(9) = lstData.List(9) & CStr(Value)
    
    Value = BaseBoard.Name
    lstData.AddItem Left("Name" & Space(35), 35)
    lstData.List(10) = lstData.List(10) & CStr(Value)

    Value = BaseBoard.OtherIdentifyingInfo
    lstData.AddItem Left("Other Identifying Info" & Space(35), 35)
    lstData.List(11) = lstData.List(11) & CStr(Value)

    Value = BaseBoard.PartNumber
    lstData.AddItem Left("Part Number" & Space(35), 35)
    lstData.List(12) = lstData.List(12) & CStr(Value)

    Value = BaseBoard.PoweredOn
    lstData.AddItem Left("Powered On" & Space(35), 35)
    lstData.List(13) = lstData.List(13) & CStr(Value)

    Value = BaseBoard.Product
    lstData.AddItem Left("Product" & Space(35), 35)
    lstData.List(14) = lstData.List(14) & CStr(Value)

    Value = BaseBoard.Removable
    lstData.AddItem Left("Removable" & Space(35), 35)
    lstData.List(15) = lstData.List(15) & CStr(Value)

    Value = BaseBoard.Replaceable
    lstData.AddItem Left("Replaceable" & Space(35), 35)
    lstData.List(16) = lstData.List(16) & CStr(Value)
    
    Value = BaseBoard.RequirementsDescription
    lstData.AddItem Left("Requirements Description" & Space(35), 35)
    lstData.List(17) = lstData.List(17) & CStr(Value)

    Value = BaseBoard.RequiresDaughterBoard
    lstData.AddItem Left("Requires Daughter Board" & Space(35), 35)
    lstData.List(18) = lstData.List(18) & CStr(Value)

    Value = BaseBoard.SerialNumber
    lstData.AddItem Left("Serial Number" & Space(35), 35)
    lstData.List(19) = lstData.List(19) & CStr(Value)

    Value = BaseBoard.SKU
    lstData.AddItem Left("SKU" & Space(35), 35)
    lstData.List(20) = lstData.List(20) & CStr(Value)

    Value = BaseBoard.SlotLayout
    lstData.AddItem Left("Slot Layout" & Space(35), 35)
    lstData.List(21) = lstData.List(21) & CStr(Value)

    Value = BaseBoard.SpecialRequirements
    lstData.AddItem Left("Special Requirements" & Space(35), 35)
    lstData.List(22) = lstData.List(22) & CStr(Value)

    Value = BaseBoard.Status
    lstData.AddItem Left("Status" & Space(35), 35)
    lstData.List(23) = lstData.List(23) & CStr(Value)

    Value = BaseBoard.Tag
    lstData.AddItem Left("Tag" & Space(35), 35)
    lstData.List(24) = lstData.List(24) & CStr(Value)

    Value = BaseBoard.Version
    lstData.AddItem Left("Version" & Space(35), 35)
    lstData.List(25) = lstData.List(25) & CStr(Value)

    Value = BaseBoard.Weight
    lstData.AddItem Left("Tag" & Space(35), 35)
    lstData.List(26) = lstData.List(26) & CStr(Value) & "pounds"

    Value = BaseBoard.Width
    lstData.AddItem Left("Width" & Space(35), 35)
    lstData.List(27) = lstData.List(27) & CStr(Value) & "inches"
    
    Me.MousePointer = vbNormal
End Sub
