VERSION 5.00
Begin VB.Form frmSelect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Video Source"
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4155
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   4155
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   300
      Left            =   3120
      TabIndex        =   2
      Top             =   720
      Width           =   900
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select"
      Default         =   -1  'True
      Height          =   300
      Left            =   2160
      TabIndex        =   1
      Top             =   720
      Width           =   900
   End
   Begin VB.ComboBox cmboSource 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   3975
   End
   Begin VB.Label lblDriver 
      Caption         =   "Installed Drivers:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSelect_Click()
   
    Dim sTitle As String
    Dim Caps As CAPDRIVERCAPS
    
    If cmboSource.ListIndex <> -1 Then
        
        '// Connect the capture window to the driver
        If capDriverConnect(lwndC, cmboSource.ListIndex) Then
    
            '// Get the capabilities of the capture driver
            capDriverGetCaps lwndC, VarPtr(Caps), Len(Caps)
            
            '// If the capture driver does not support a dialog, grey it out
            '// in the menu bar.
            frmMain.mnuSource.Enabled = Caps.fHasDlgVideoSource
            frmMain.mnuFormat.Enabled = Caps.fHasDlgVideoFormat
            frmMain.mnuDisplay.Enabled = Caps.fHasDlgVideoDisplay
        
            sTitle = cmboSource.Text
            
            SetWindowText lwndC, sTitle
            ResizeCaptureWindow lwndC
        End If
    
    End If
    
    
    Unload Me
   
End Sub


Private Sub Form_Load()
    
    Dim lpszName As String * 100
    Dim lpszVer As String * 100
    Dim x As Integer
    Dim lResult As Long
    Dim Caps As CAPDRIVERCAPS
    
    '// Get a list of all the installed drivers
    x = 0
    Do
        lResult = capGetDriverDescriptionA(x, lpszName, 100, lpszVer, 100)   '// Retrieves driver info
        If lResult Then
            cmboSource.AddItem lpszName
            x = x + 1
        End If
    Loop Until lResult = False

    '// Get the capabilities of the current capture driver
    lResult = capDriverGetCaps(lwndC, VarPtr(Caps), Len(Caps))
    
    '// Select the driver that is currently being used
    If lResult Then cmboSource.ListIndex = Caps.wDeviceIndex

End Sub


