VERSION 5.00
Begin VB.Form frmMainAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4905
   Icon            =   "frmMainAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMainAbout.frx":000C
   ScaleHeight     =   3075
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblPage 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "http://free.prohosting.com/~thelung/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H001212A0&
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label lblPageLabel 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Web Page"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H001212A0&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblContactLabel 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Contact"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H001212A0&
      Height          =   255
      Left            =   2880
      TabIndex        =   6
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label lblContact 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "the_lung_@hotmail.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H001212A0&
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label lblStarted 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "March 30, 2000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H001212A0&
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label lblMadeBy 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Professor Lung"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H001212A0&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label lblStartedLabel 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Started on"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H001212A0&
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lblMadeByLabel 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Made by"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H001212A0&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   1095
   End
End
Attribute VB_Name = "frmMainAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Unload(Cancel As Integer)
    Unload Me 'Just unloads this window not the whole program
End Sub
