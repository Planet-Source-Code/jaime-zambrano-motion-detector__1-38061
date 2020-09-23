VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "VB Memcap"
   ClientHeight    =   4860
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5760
   Icon            =   "VBmemcapNL.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   324
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   384
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   1560
      Top             =   1320
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   120
      ScaleHeight     =   1515
      ScaleWidth      =   2235
      TabIndex        =   1
      Top             =   2520
      Width           =   2295
   End
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   4605
      Width           =   5760
      _ExtentX        =   10160
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Do not Change Image Size with dragging cursor"
      Height          =   735
      Left            =   2520
      TabIndex        =   2
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuAllocate 
         Caption         =   "&Allocate"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu mnuControl 
      Caption         =   "&Control"
      Begin VB.Menu mnuStart 
         Caption         =   "&Start"
      End
      Begin VB.Menu mnuDisplay 
         Caption         =   "&Display"
      End
      Begin VB.Menu mnuFormat 
         Caption         =   "&Format"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuSource 
         Caption         =   "S&ource"
      End
      Begin VB.Menu mnuCompression 
         Caption         =   "Co&mpression"
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelect 
         Caption         =   "&Select"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuScale 
         Caption         =   "Sc&ale"
         Checked         =   -1  'True
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuPreview 
         Caption         =   "&Preview"
         Checked         =   -1  'True
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAlwaysVisible 
         Caption         =   "Al&ways Visible"
         Shortcut        =   ^W
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
    '----------------------------
    Dim StdVar As Double
    Dim Promedio As Double
    Dim Suma As Double
    Dim CuentaIMG As Integer
    Dim CuentaCiclo(10) As Integer
    Dim SumaCuadr As Double
     
    Dim Matriz_Prom(10, 10) As Double
    Dim Matriz_Desvest(10, 10) As Double
     
    Dim Prom_Cadena(10) As Double
    Dim Prom_Cadena_Anterior(10) As Double
    Dim Desvest_Cadena(10) As Double
    Dim Umbral_Prom As Double
    Dim Umbral_Desvest As Double
    Dim SeHallenadoCadena(10) As Boolean
    Dim SeIncorporaNuePixel(10) As Boolean
    '----------------------------
    '----------------------------
    Dim lPicLeft  As Single 'Long
    Dim lPicTop As Single 'Long
     
    Dim lMoveX As Long
    Dim lMoveY As Long
    Dim bForwardMove As Boolean
    Dim bDownMove As Boolean
    Dim bMoving As Boolean
     
    Dim lTransColor As Long
    '----------------------------
Const Ancho = 164
Const Alto = 124
Const Delta = 1000000
Dim Sheet1 As Object
    
Dim ColorPunto As Long
Dim MatrizPunto(1000) As Long
Dim SumaPunto(10000) As Long

Dim Locked As Boolean

Dim i As Integer
Dim j As Integer

Dim UmbralTrFalse As Double
Private C As ColorComponent

Private Sub Form_Load()

    Locked = False

    Picture1.AutoSize = True
    Dim lpszName As String * 100
    Dim lpszVer As String * 100
    Dim Caps As CAPDRIVERCAPS
     '-----------------------------
    Timer1.Interval = BuscaValor(3) 'In the file "InfoPlanilla.txt"
                                  'change the third value to
                                  'increase or decrease surveillance
                                  ' cycles in frequency.
     '-----------------------------
    Umbral_Prom = BuscaValor(1)
    Umbral_Desvest = BuscaValor(2)
    UmbralTrFalse = BuscaValor(4) 'In the file "InfoPlanilla.txt"
                                  'change the fourth value to
                                  'increase or decrease sensibility to
                                  ' motion elements.
     '-----------------------------
     '-----------------------------
    '//Create Capture Window
    capGetDriverDescriptionA 0, lpszName, 100, lpszVer, 100  '// Retrieves driver info
    lwndC = capCreateCaptureWindowA(lpszName, WS_CAPTION Or WS_THICKFRAME Or WS_VISIBLE Or WS_CHILD, 0, 0, 160, 120, Me.hwnd, 0)

    '// Set title of window to name of driver
    SetWindowText lwndC, lpszName
     
    '// Set the video stream callback function
    capSetCallbackOnStatus lwndC, AddressOf MyStatusCallback
    capSetCallbackOnError lwndC, AddressOf MyErrorCallback
     
    '// Connect the capture window to the driver
    If capDriverConnect(lwndC, 0) Then

        capDriverGetCaps lwndC, VarPtr(Caps), Len(Caps)
         

        If Caps.fHasDlgVideoSource = 0 Then mnuSource.Enabled = False
        If Caps.fHasDlgVideoFormat = 0 Then mnuFormat.Enabled = False
        If Caps.fHasDlgVideoDisplay = 0 Then mnuDisplay.Enabled = False
         

        capPreviewScale lwndC, True
             
        '// preview rate in ms
        capPreviewRate lwndC, 66
         
        '// previewing image from Webcam
        capPreview lwndC, True
             
       
        ResizeCaptureWindow lwndC

    End If

    Picture1.ScaleMode = vbPixels
    Picture1.AutoRedraw = True
    lPicLeft = 1
    lPicTop = 1
MakeTopMost (hwnd)

End Sub

Private Sub Form_Unload(Cancel As Integer)

    '// Disable all callbacks
    capSetCallbackOnVideoStream lwndC, vbNull
    capSetCallbackOnWaveStream lwndC, vbNull
    capSetCallbackOnCapControl lwndC, vbNull
    capSetCallbackOnError lwndC, vbNull
    capSetCallbackOnStatus lwndC, vbNull
    capSetCallbackOnYield lwndC, vbNull
    capSetCallbackOnFrame lwndC, vbNull
    

End Sub

Private Sub mnuAllocate_Click()

 Dim sFile As String * 250
 Dim lSize As Long
  
 '// Setup swap file for capture
 lSize = 1000000
 sFile = "C:\TEMP.AVI"
 capFileSetCaptureFile lwndC, sFile
 capFileAlloc lwndC, lSize
  
End Sub

Private Sub mnuAlwaysVisible_Click()
     
    mnuAlwaysVisible.Checked = Not (mnuAlwaysVisible.Checked)
     
    If mnuAlwaysVisible.Checked Then
        SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
    Else
        SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
    End If


End Sub

Private Sub mnuCompression_Click()

     
    capDlgVideoCompression lwndC

End Sub

Private Sub mnuCopy_Click()

    capEditCopy lwndC
         
End Sub

Private Sub mnuDisplay_Click()

    capDlgVideoDisplay lwndC
     
End Sub

Private Sub mnuExit_Click()

    Unload Me
     
End Sub

Private Sub mnuFormat_Click()
    capDlgVideoFormat lwndC
    ResizeCaptureWindow lwndC

End Sub

Private Sub mnuPreview_Click()

    frmMain.StatusBar.SimpleText = vbNullString
    mnuPreview.Checked = Not (mnuPreview.Checked)
    capPreview lwndC, mnuPreview.Checked
     
End Sub

Private Sub mnuScale_Click()
     
    mnuScale.Checked = Not (mnuScale.Checked)
    capPreviewScale lwndC, mnuScale.Checked
     
    If mnuScale.Checked Then
       SetWindowLong lwndC, GWL_STYLE, WS_THICKFRAME Or WS_CAPTION Or WS_VISIBLE Or WS_CHILD
    Else
       SetWindowLong lwndC, GWL_STYLE, WS_BORDER Or WS_CAPTION Or WS_VISIBLE Or WS_CHILD
    End If

    ResizeCaptureWindow lwndC
     
End Sub

Private Sub mnuSelect_Click()
     
    frmSelect.Show vbModal, Me

End Sub

Private Sub mnuSource_Click()
    capDlgVideoSource lwndC

End Sub

Private Sub mnuStart_Click()

    Dim sFileName As String
    Dim CAP_PARAMS As CAPTUREPARMS
     
    capCaptureGetSetup lwndC, VarPtr(CAP_PARAMS), Len(CAP_PARAMS)
     
    CAP_PARAMS.dwRequestMicroSecPerFrame = (1 * (10 ^ 6)) / 30  ' 30 Frames per second
    CAP_PARAMS.fMakeUserHitOKToCapture = True
    CAP_PARAMS.fCaptureAudio = False
     
    capCaptureSetSetup lwndC, VarPtr(CAP_PARAMS), Len(CAP_PARAMS)
     
    sFileName = "C:\myvideo.avi"
     
    capCaptureSequence lwndC
    capFileSaveAs lwndC, sFileName

End Sub
Private Sub StatusBar1_PanelClick(ByVal Panel As ComctlLib.Panel)

End Sub
Private Sub CargadePantalla()
    Picture1.Picture = Clipboard.GetData(vbCFDIB)
End Sub
Private Function ValoresLinea(Xa As Integer, Ya As Integer, Xb As Integer, Yb As Integer) As ParPromDest
'Yields Avg and StdVar for the specified line
Dim i As Single, j As Single, ve As Double, Suma As Double, SumaCuadr As Double
IniPunto = 1

If Yb - Ya > Xb - Xa Then
    'Analysis through Y axis
    ve = (Xb - Xa) / (Yb - Ya)
    For j = Ya To Yb
    i = Int((j - Ya) * ve) + Xa
  
    GetColores Picture1, i, j, C
   
            Suma = Suma + C.SumaColInt
            SumaCuadr = SumaCuadr + (C.SumaColInt ^ 2)
    Next j
             
            ValoresLinea.Promedio = Int(Suma / (Yb - Ya + 1))
            ValoresLinea.StdVar = Int(Sqr(((Yb - Ya + 1) * SumaCuadr - (Suma ^ 2)) / (Yb - Ya + 1) / (Yb - Ya)))
           
Else
    'Analysis through X axis
    ve = (Yb - Ya) / (Xb - Xa)
    For i = Xa To Xb
    j = Int((i - Xa) * ve) + Ya
  
    GetColores Picture1, i, j, C
   
            Suma = Suma + C.SumaColInt
            SumaCuadr = SumaCuadr + (C.SumaColInt ^ 2)
    Next i

            ValoresLinea.Promedio = Int(Suma / (Xb - Xa + 1))
            ValoresLinea.StdVar = Int(Sqr(((Xb - Xa + 1) * SumaCuadr - (Suma ^ 2)) / (Xb - Xa + 1) / (Xb - Xa)))
End If
End Function

Private Sub Timer1_Timer()
Dim DiagoPrinc As ParPromDest
Dim Signal1 As Boolean
'-----*-*-*-*-*-*-*-*-*
Static DiagoPromAnte As Double
Static Diferencia As Double
Static DiferenciaAnte As Double
Static SegDerivada As Double
Static Contador As Integer
'-----*-*-*-*-*-*-*-*-*
If Locked = False Then
    Pausa (10)
    Locked = True
End If
    Signal1 = False
    capEditCopy lwndC: CargadePantalla

    
'---*-*-*-*-*-*'Routine that counts each "Chain" Cycle
DiagoPrinc = CicleaLinea(1, 58, 159, 60, 1)
'---*-*-*-*-*-*

If DiagoPromAnte <> 0 Then
        Diferencia = Abs(DiagoPromAnte - DiagoPrinc.Promedio)
    If Diferencia > UmbralTrFalse Then
        DiagoPromAnte = DiagoPrinc.Promedio
        Signal1 = True
        
    Else
        DiagoPromAnte = DiagoPrinc.Promedio
    End If
Else
    DiagoPromAnte = DiagoPrinc.Promedio
End If
' the second derivative allows to make the difference
' between an out-of Avg. image and the normal lightning persitence
' the Avg. value has to overcome when restoring to usual image
' conditions
SegDerivada = Diferencia - DiferenciaAnte
    If SegDerivada > 0 Then
    DiferenciaAnte = Diferencia
        If Signal1 = True Then
            Contador = Contador + 1
            Debug.Print "Saving", Contador
            SavePicture Picture1.Image, App.Path & "\ImageN" & Contador & ".BMP"
        End If
    Else
    DiferenciaAnte = Diferencia
    End If

End Sub
Private Function CicleaLinea(Xa As Integer, Ya As Integer, Xb As Integer, Yb As Integer, NL As Integer) As ParPromDest
Dim Resultado As ParPromDest
Resultado = ValoresLinea(Xa, Ya, Xb, Yb)


If CuentaCiclo(NL) = 9 Then ' Cycle n°9 already executed '
    CuentaCiclo(NL) = 0
    SeHallenadoCadena(NL) = True
End If

CuentaCiclo(NL) = CuentaCiclo(NL) + 1

'---*-*-*-*-*-*    '---*-*-*-*-*-*    '---*-*-*-*-*-*
If CuentaCiclo(NL) = 1 Then  'First Value added to "Chain"
       If SeHallenadoCadena(NL) = False Then
                        Matriz_Desvest(1, NL) = Resultado.StdVar
                        Matriz_Prom(1, NL) = Resultado.Promedio
                        SeIncorporaNuePixel(NL) = True
       Else
                    SeIncorporaNuePixel(NL) = False
        '---*-*-*-*-*-* ' Condition for adding latest Value to "Chain"
                    If Abs(Matriz_Desvest(9, NL) - Resultado.StdVar) < Umbral_Desvest And Abs(Matriz_Prom(9, NL) - Resultado.Promedio) < Umbral_Prom Then
                        Matriz_Desvest(CuentaCiclo(NL), NL) = Resultado.StdVar
                        Matriz_Prom(CuentaCiclo(NL), NL) = Resultado.Promedio
                        SeIncorporaNuePixel(NL) = True
                    End If
       '---*-*-*-*-*-*-*-*-*-*-*-*
          PROMEDIAR_CADENA (NL)
        '---*-*-*-*-*-*-* ' Condition for handling Image saving
                    If Abs(Desvest_Cadena(NL) - Resultado.StdVar) > Umbral_Desvest And Abs(Prom_Cadena(NL) - Resultado.Promedio) > Umbral_Prom Then
                    CuentaIMG = CuentaIMG + 1
                    
                    End If
       End If

     
Else ' CuentaCiclo(NL) >= 2 Then
        '---*-*-*-*-*-*
        SeIncorporaNuePixel(NL) = False

        If Abs(Matriz_Desvest(CuentaCiclo(NL) - 1, NL) - Resultado.StdVar) < Umbral_Desvest And Abs(Matriz_Prom(CuentaCiclo(NL) - 1, NL) - Resultado.Promedio) < Umbral_Prom Then
                    
          Matriz_Desvest(CuentaCiclo(NL), NL) = Resultado.StdVar
            Matriz_Prom(CuentaCiclo(NL), NL) = Resultado.Promedio
            SeIncorporaNuePixel(NL) = True
        End If

          PROMEDIAR_CADENA (NL)

        If Abs(Desvest_Cadena(NL) - Resultado.StdVar) > Umbral_Desvest And Abs(Prom_Cadena(NL) - Resultado.Promedio) > Umbral_Prom Then

            CuentaIMG = CuentaIMG + 1
            
        End If
         
End If

    If SeIncorporaNuePixel(NL) = False Then
        CuentaCiclo(NL) = CuentaCiclo(NL) - 1
    End If

CicleaLinea.Promedio = Prom_Cadena(NL)
CicleaLinea.StdVar = Desvest_Cadena(NL)

End Function

Private Sub PROMEDIAR_CADENA(NL As Integer)
Dim j As Integer
If SeIncorporaNuePixel(NL) = True Then

    If SeHallenadoCadena(NL) = True Then
        Prom_Cadena(NL) = (Prom_Cadena(NL) * 8 + Matriz_Prom(CuentaCiclo(NL), NL)) / 9
        Desvest_Cadena(NL) = (Desvest_Cadena(NL) * 8 + Matriz_Desvest(CuentaCiclo(NL), NL)) / 9
     
    Else
        Prom_Cadena(NL) = 0: Desvest_Cadena(NL) = 0
        For j = 1 To 9
                If Matriz_Prom(j, NL) = 0 And Matriz_Desvest(j, NL) = 0 Then Exit For
                Prom_Cadena(NL) = Prom_Cadena(NL) + Matriz_Prom(j, NL)
                Desvest_Cadena(NL) = Desvest_Cadena(NL) + Matriz_Desvest(j, NL)
        Next j
        If j <> 1 Then
        Prom_Cadena(NL) = Prom_Cadena(NL) / (j - 1)
        Desvest_Cadena(NL) = Desvest_Cadena(NL) / (j - 1)
        End If
    End If

End If ' No new pixel added so there is
' No need for calculating new avg.
End Sub

Private Sub Pausa(TimePause As Variant)
Dim Inicio
 ''----------------------------------------------
        Inicio = Timer   ' Sets starting time for pause (in sec.)
            Do While Timer < Inicio + TimePause
                DoEvents   'Shifts to other process
                Loop
End Sub

 Private Function BuscaValor(i As Integer) As String
        Dim LeeString, cuenta
            Open App.Path & "\InfoPlanilla.txt" For Input As #1
                 For cuenta = 1 To i
            Input #1, LeeString
                 Next cuenta
        Close #1

 BuscaValor = LeeString
  
 End Function

