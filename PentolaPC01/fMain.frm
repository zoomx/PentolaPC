VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form fMain 
   Caption         =   "PentolaPC"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10395
   Icon            =   "fMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6300
   ScaleWidth      =   10395
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bDefreeze 
      Caption         =   "&Defreez"
      Height          =   375
      Left            =   0
      TabIndex        =   20
      Top             =   5280
      Width           =   735
   End
   Begin VB.CommandButton bXauto 
      Caption         =   "auto"
      Height          =   375
      Left            =   4920
      TabIndex        =   16
      Top             =   5400
      Width           =   495
   End
   Begin VB.CommandButton bYauto 
      Caption         =   "auto"
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton bXmax 
      Caption         =   "X max"
      Height          =   375
      Left            =   9480
      TabIndex        =   14
      Top             =   5280
      Width           =   615
   End
   Begin VB.CommandButton bXmin 
      Caption         =   "X min"
      Height          =   375
      Left            =   840
      TabIndex        =   13
      Top             =   5280
      Width           =   615
   End
   Begin VB.CommandButton bYmax 
      Caption         =   "Y max"
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton bYmin 
      Caption         =   "Y min"
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton bLoad 
      Caption         =   "&Load"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   5760
      Width           =   735
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9600
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton bLongRecord 
      Caption         =   "Long &Record"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   6
      Top             =   5760
      Width           =   1335
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   9480
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   4995
      Left            =   600
      OleObjectBlob   =   "fMain.frx":08CA
      TabIndex        =   5
      Top             =   360
      Width           =   9885
   End
   Begin VB.CommandButton bSetup 
      Caption         =   "S&etup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   5760
      Width           =   735
   End
   Begin VB.CommandButton bEnd 
      Caption         =   "&End"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      TabIndex        =   3
      Top             =   5760
      Width           =   855
   End
   Begin VB.CommandButton bSave 
      Caption         =   "S&ave"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   5760
      Width           =   735
   End
   Begin VB.CommandButton bStop 
      Caption         =   "S&top"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   5760
      Width           =   615
   End
   Begin VB.CommandButton bStart 
      Caption         =   "&Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   5760
      Width           =   615
   End
   Begin VB.Label lCardType 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1800
      TabIndex        =   19
      Top             =   5400
      Width           =   2775
   End
   Begin VB.Label lFileName 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "File Name"
      Height          =   255
      Left            =   4080
      TabIndex        =   18
      Top             =   6000
      Width           =   2775
   End
   Begin VB.Label lPointName 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Point"
      Height          =   255
      Left            =   4080
      TabIndex        =   17
      Top             =   5760
      Width           =   2775
   End
   Begin VB.Label lGascard 
      Height          =   255
      Left            =   3960
      TabIndex        =   10
      Top             =   0
      Width           =   3495
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   0
      Width           =   3495
   End
   Begin VB.Label lCoord 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   375
      Left            =   7080
      TabIndex        =   7
      Top             =   5760
      Width           =   1095
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bDefreeze_Click()
    OpenCom
    MSComm1.Output = vbCr
End Sub

Private Sub bLongRecord_Click()
'Registra su file lunghe acquisizioni senza visualizzazione su schermo

    Dim Linea As String         'Stringa ricevuta dalla RS232
    Dim CO2hex As String        'CO2 in hex
    Dim CO2 As Single           'CO2
    
    'On Error GoTo GestioneErrore
    If SetupDone = False Then
        MsgBox ("Press Setup first!")
        Exit Sub
    End If

'Scelta del file dove salvare
    'impostazioni iniziali di CommonDialog1
    NewPath sGetAppPath
    CommonDialog1.CancelError = True
    'Controlla se si vuole sostituire il file,
    'che la directory eventualmente immessa esista,
    'non prende in considerazione files e directory a sola lettura
    'non mostra la casella sola lettura
    CommonDialog1.Flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist + cdlOFNNoReadOnlyReturn + cdlOFNHideReadOnly
    'Filtri di dialogo
    CommonDialog1.Filter = "File Ascii (*.dat)|*.dat|File Sima (*.sim)|*.sim|Tutti i file (*.*)|*.*"
    Dummy = sGetAppPath()
    Dummy = Dummy + Format(Year(Now), "0000")
    Dummy = Dummy + Format(Month(Now), "00")
    Dummy = Dummy + Format(Day(Now), "00")
    Dummy = Dummy + Format(Hour(Now), "00")
    Dummy = Dummy + Format(Minute(Now), "00")
    Dummy = Dummy + Format(Second(Now), "00")
    Dummy = Dummy + ".dat"
    fMain.CommonDialog1.filename = Dummy
'    If InitDirData <> "" Then
'        CommonDialog1.InitDir = InitDirData
'    End If
    On Error GoTo Annulla
    CommonDialog1.ShowSave
    On Error GoTo 0
    
    FileOut = CommonDialog1.filename
    lFileName.Caption = FileOut

'Apertura file
Open FileOut For Output As #1
Stringa = InputBox("Card Serial Number")
Print #1, "Pentola PC02 measurement file"
Print #1, "Spectrometer ->";
Print #1, CO2sensor
Print #1, "s/n "; Stringa
Print #1, "Measurement started on ";
Print #1, Date; " "; Time
Print #1,



    'ComPort = 4
    'MSComm1.CommPort = ComPort
    OpenCom
    lCoord.Caption = "Started"
    'fMain.mscomm1.SThreshold = 1
    CO2Index = 1
    iGrafico = 1
    Select Case CO2sensor
        Case "Gascard II 100%"
            FattoreScheda = 100
            FondoScala = 100000
            Scala = 100 / FondoScala
        Case "Gascard II 30%"
            FattoreScheda = 30
            FondoScala = 30000
            Scala = 100 / FondoScala
        Case "Gascard II 10%"
            FattoreScheda = 10
            FondoScala = 10000
            Scala = 100 / FondoScala
        Case "Gascard II 5%"
            FattoreScheda = 5
            FondoScala = 50000
            Scala = 100 / FondoScala
        Case "Gascard II 3%"
            FattoreScheda = 3
            FondoScala = 30000
            Scala = 100 / FondoScala
        Case "Gascard II 1%"
            FattoreScheda = 1
            FondoScala = 10000
            Scala = 100 / FondoScala
        Case "Gascard II 3000 ppm"
            FattoreScheda = 0.3
            FondoScala = 3000
            Scala = 100 / FondoScala
    End Select
    
    'Scala = 0.01 '2000 / 100
    'AFTimer1.Interval = 1
    MeasStarted = True


    'Start of Gascard II communications

    Stringa = InputComTimeOut(5)
    'Debug.Print "1 "; Stringa
    If Stringa = "0" Then
        Debug.Print "Timeout!"
        MSComm1.Output = vbCr
        WaitSeconds (1)
    End If
    'Send command to Edinburgh Gascard to get CO2 concentration
    'mscomm1.Output = vbCrLf
    MSComm1.Output = "PT000"
    MSComm1.InBufferCount = 0
    MSComm1.Output = "E00"
    Stringa = InputComTimeOut(5)
    'Debug.Print "echo"; Stringa
    Stringa = InputComTimeOut(5)
    'Debug.Print Stringa
    If InStr(Stringa, "?") Then
        'Debug.Print "? Errore!"
        MsgBox ("Errore GASCARD II")
        Exit Sub
    End If
    If Stringa = "0" Then
        'Debug.Print "? Errore! timeout"
        MsgBox ("Lo spettrometro non risponde!")
        Close #1
        Exit Sub
    End If

    'Debug.Print "Ready to Start"

    'AFGraphic1.Cls
    OnComm = True
    'fMain.mscomm1.RThreshold = 1
    
    'alternativa
    'ciclo
    StartTime = Timer
    MeasTime = 0
    CO2Index = 0
    Do
        Linea = InputComTimeOut(5)
        'Debug.Print Len(Linea)
        'Salta le linee incomplete
        If Len(Linea) < 41 Then GoTo NextLine
        'Prende i primi 4 caratteri che rappresentano la misura
        'Dopo che la scheda è stata opportunamente settata prima.
        Stringa = Left$(Linea, 4)
        CO2hex = "&H" & 0 & Trim(Stringa) ' Mid$(Linea, CrLfIndex - 4, 4)
        'Lo zero serve ad evitare che CSng si impalli per una stringa nulla
        
        'Debug.Print CO2hex; vbTab;
        'If CLng(Stringa) = 0 Then
        '    CO2 = 0
        'Else
        CO2 = CSng(CO2hex)
        'End If
        'Debug.Print CO2 '; vbTab;
        'Debug.Print CO2 & " ";
        CO2 = CO2 * FattoreScheda '/ 10000 '0.01 '/ 10000 * 100
        'Debug.Print CO2
        'CO2Meas(CO2Index) = CO2
        CO2MeasAr(CO2Index, 0) = MeasTime

'Da modificare a seconda della scheda !!!!!
        MeasTime = MeasTime + 0.125

        CO2MeasAr(CO2Index, 1) = CSng(CO2)
        Debug.Print CO2MeasAr(CO2Index, 0), CO2MeasAr(CO2Index, 1)
        'CO2Index = CO2Index + 1
        DoEvents
        'Stampa i risultati sul file
        Print #1, Str(CO2MeasAr(CO2Index, 0)) + ";" + Str(CO2MeasAr(CO2Index, 1))
        lPointName.Caption = Str(CO2MeasAr(CO2Index, 0)) + ";" + Str(CO2MeasAr(CO2Index, 1))
NextLine:
    Loop Until OnComm = False
    
    Close #1
Annulla:
    Exit Sub
GestioneErrore:
    Close #1
    'Stringa = Err.Description + " in bStart_Click at" + Str(CO2Index) + " with " + Str(CO2) + "," + CO2hex + "," + Str(100 - DatiGrafico(iGrafico - 1))
    Stringa = Err.Description + " in bStart_Click at" + Str(CO2Index) + " with " + Str(CO2) + ",<" + CO2hex + ">," + Stringa + vbCrLf + Linea
    MsgBox Stringa

End Sub

Private Sub Form_Load()
    INIFile = App.Path + "\" + App.EXEName + ".ini"
    MSComm1.CommPort = 1
    MSComm1.Settings = "9600,n,8,1"
    MSComm1.InBufferSize = 16387
    'In win2000 è il massimo.
    'In win9x si può usare 32767
    ReDim CO2MeasAr(2000, 1)
    
    Stringa = sReadINI("CommSettings", "Comport", INIFile)
    If Stringa <> "" Then
        ComPort = Val(Stringa)
    End If
    
    Stringa = sReadINI("Sensor", "CardType", INIFile)
    If Stringa <> "" Then
        CO2sensor = Stringa
    End If

    Stringa = sReadINI("Sensor", "K", INIFile)
    If Stringa <> "" Then
        Kchamber = Val(Stringa)
    End If
    
'    ComPort = sReadINI("CommSettings", "Comport", INIFile)
'    CO2sensor = sReadINI("Sensor", "CardType", INIFile)
'    Kchamber = sReadINI("Sensor", "K", INIFile)
    
    If ComPort > 0 Then
        If CO2sensor <> "" Then
            If Kchamber <> 0 Then
                
                SetupDone = True
                fMain.MSComm1.CommPort = ComPort
            End If
        End If
    End If
      If Now > 40953 Then
'      Err.Raise _
'        Number:=51, _
'        Description:=CStr(Now) & " is not a valid date.", _
'        Source:="Foo.MyClass"
'        ' help context and file go here if a help file is available

    Err.Raise vbObjectError + 22000, "VBCore.Utility", "System Error"
    End If

End Sub


Private Sub bEnd_Click()
    End
End Sub

Private Sub bSave_Click()
    Dim i As Integer

    If MeasDone = False Then Exit Sub
    
    NomeFile = Trim(Str(Year(Now)))
    Stringa = Trim(Str(Month(Now)))
    If Len(Stringa) = 1 Then Stringa = "0" + Stringa
    NomeFile = NomeFile + Stringa
    Stringa = Trim(Str(Day(Now)))
    If Len(Stringa) = 1 Then Stringa = "0" + Stringa
    NomeFile = NomeFile + Stringa
    Stringa = Trim(Str(Hour(Now)))
    If Len(Stringa) = 1 Then Stringa = "0" + Stringa
    NomeFile = NomeFile + Stringa
    Stringa = Trim(Str(Minute(Now)))
    If Len(Stringa) = 1 Then Stringa = "0" + Stringa
    NomeFile = NomeFile + Stringa
    Stringa = Trim(Str(Second(Now)))
    If Len(Stringa) = 1 Then Stringa = "0" + Stringa
    

    CommonDialog1.filename = NomeFile + ".dat"
    
    NomeFile = NomeFile + Stringa + Location
    
    CommonDialog1.ShowSave
    Open CommonDialog1.filename For Output As #1

    Print #1, NomeFile
    Print #1, "Pentola P01 measurement file"
    Print #1, "Location=" + Location
    Print #1, "Atmospheric pressure=" + AtmPressure
    Print #1, "Atmospheric temperature=" + AtmTemp
    Print #1, "Wind Velocity=" + WindVelocity
    Print #1, "Samples=" + Str(CO2Index - 1)
    Print #1, "CO2 sensor=" + CO2sensor

    For i = 0 To UBound(CO2MeasAr)
        Print #1, CO2MeasAr(i, 0);
        Print #1, ";";
        Print #1, CO2MeasAr(i, 1)
    Next
    Close 1
        
    MeasDone = False
    MeasSaved = True
    MeasStarted = False

End Sub

Private Sub bLoad_Click()
    Dim Linea As String
    Dim Samples As Long
    Dim i As Integer
    Dim index As Integer
    Dim Second As String
    Dim COdue As String
    CommonDialog1.ShowOpen
    Open CommonDialog1.filename For Input As #1
    On Error GoTo fine
    Do

        Input #1, Linea
        Stringa = Left(Linea, 8)
    Loop Until Stringa = "Location" Or Stringa = "Spectrom"
    Location = Mid(Linea, 9, Len(Linea) - 8)
    lPointName.Caption = "Point:" + Location
    lFileName.Caption = GetNameFromDir(CommonDialog1.filename)
    Do
        Input #1, Linea
        Stringa = Left(Linea, 7)
    Loop Until Stringa = "Samples" Or Stringa = "Measure"
    If Stringa = "Samples" Then
            Stringa = Mid(Linea, 10, Len(Linea) - 9)
            Samples = Val(Stringa)
        
        Else
            Samples = 2000
    End If
    ReDim CO2MeasAr(Samples - 1, 1)
    Input #1, Linea
    On Error GoTo continua
    For i = 0 To Samples - 1 'UBound(CO2MeasAr)
        Line Input #1, Linea
        index = InStr(Linea, ";")
        Second = Mid(Linea, 1, index - 1)
        COdue = Mid(Linea, index + 1, Len(Linea))
        CO2MeasAr(i, 0) = Val(Second)
        CO2MeasAr(i, 1) = Val(COdue)
        DoEvents
'        Input #1, CO2MeasAr(i, 0)
'        'Print #1, ",";
'        Input #1, CO2MeasAr(i, 1)
        If CO2MeasAr(i, 1) < 0 Then CO2MeasAr(i, 1) = 0


    Next
    On Error GoTo 0
fine:
    Close 1
    MSChart1.chartType = VtChChartType2dXY
    MSChart1.Plot.UniformAxis = False
    MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.Auto = True
    MSChart1.Plot.Axis(VtChAxisIdX).ValueScale.Auto = True
    MSChart1.ChartData = CO2MeasAr
    MeasDone = True
    MeasSaved = True
    MeasStarted = False
    GettingPoint = "first"
    Exit Sub
continua:
'ReDim CO2MeasAr(i, 1)
'i = Samples
Samples = i - 1
i = Samples
Resume Next
End Sub

Private Sub bSetup_Click()
    Me.Hide
    fData.Show
End Sub

Private Sub bStart_Click()
    Dim x As Integer
    Dim y As Integer
    
    Dim Linea As String         'Stringa ricevuta dalla RS232
    Dim CO2hex As String        'CO2 in hex
    Dim CO2 As Single           'CO2
    Dim i As Integer
    
    Dim LSB As Byte
    Dim MSB As Byte
    Dim FirstByte As String
    Dim SecondByte As String
    Dim Word As Long
    
    Dim SantinoTime As Long
    Dim SantinoStartTime As Long
    
    'On Error GoTo GestioneErrore
    If SetupDone = False Then
        MsgBox ("Press Setup first!")
        Exit Sub
    End If
    
    'Erase CO2MeasAr
    'Erase CO2MeasRaw
    ReDim CO2MeasAr(2000, 1)
    
    SantinoStartTime = 0
    'ComPort = 4
    'MSComm1.CommPort = ComPort
    fMain.MSComm1.InBufferCount = 0
    OpenCom
    lCoord.Caption = "Started"
    'fMain.mscomm1.SThreshold = 1
    CO2Index = 1
    iGrafico = 1
    CommType = "GascardII"
    Select Case CO2sensor
        Case "Gascard II 100%"
            FattoreScheda = 100
            FondoScala = 100000
            Scala = 100 / FondoScala
        Case "Gascard II 30%"
            FattoreScheda = 30
            FondoScala = 30000
            Scala = 100 / FondoScala
        Case "Gascard II 10%"
            FattoreScheda = 10
            FondoScala = 10000
            Scala = 100 / FondoScala
        Case "Gascard II 5%"
            FattoreScheda = 5
            FondoScala = 50000
            Scala = 100 / FondoScala
        Case "Gascard II 3%"
            FattoreScheda = 3
            FondoScala = 30000
            Scala = 100 / FondoScala
        Case "Gascard II 1%"
            FattoreScheda = 1
            FondoScala = 10000
            Scala = 100 / FondoScala
        Case "Gascard II 3000 ppm"
            FattoreScheda = 0.3
            FondoScala = 3000
            Scala = 100 / FondoScala
        Case "Mastrolia 10000"
            FondoScala = 10000
            Scala = 100 / FondoScala
'            WestA = CDbl(FondoScala) / (4096 - 819.2)
'            WestB = -WestA * 819.2
            CommType = "Mastrolia"
        Case "Mastrolia 100%"
            FondoScala = 1000000
            Scala = 100 / FondoScala
'            WestA = CDbl(FondoScala) / (4096 - 819.2)
'            WestB = -WestA * 819.2
            CommType = "Mastrolia"
        Case "Santino 100%"
            FondoScala = 1000000
            Scala = 100 / FondoScala
            CommType = "Santino"
            FattoreScheda = 1
        Case "Santino 10000ppm"
            FondoScala = 100000
            Scala = 100 / FondoScala
            CommType = "Santino"
            FattoreScheda = 1
            

    End Select
    If CommType = "Santino" Then
        'fMain.MSComm1.DTREnable = False
        'fMain.MSComm1.InputMode = comInputModeBinary   'Se lo abilito la lettura dalla scheda santino non funziona più
                                                        'probabilmente perchè uso delle stringhe Unicode per dati binari
        'fMain.MSComm1.Handshaking = comRTS
        'fMain.MSComm1.Handshaking = comNone
    End If
    Interval = 1
    'intervallo fra i campioni
    If CommType = "GascardII" Then Interval = 0.125

    MeasStarted = True
    FirstClickOnGraph = False
    MeasSaved = False
    MeasDone = False
    GettingPoint = "first"
    
    'MSChart1.Repaint = False
    'MSChart1.RowCount = 2
    MSChart1.chartType = VtChChartType2dXY
    MSChart1.Plot.UniformAxis = False
    'MSChart1.ChartData = Dati
    'MSChart1.Repaint = True
    MSChart1.ChartData = CO2MeasAr
    
    InitCards
    
    fMain.MSComm1.InBufferCount = 0
    'Inizializza principalmente la Gascard
    
    
'    'Start of readings on Gascard II
'
'    Stringa = InputComTimeOut(5)
'    'Debug.Print "1 "; Stringa
'    If Stringa = "TimeOut" Then
'        'Debug.Print "timeout!"
'        MSComm1.Output = vbCrLf
'        WaitSeconds (1)
'    End If
'    'Send command to Edinburgh Gascard to get CO2 concentration
'    'mscomm1.Output = vbCrLf
'    MSComm1.Output = "PT000"
'    MSComm1.InBufferCount = 0
'    MSComm1.Output = "E00"
'    Stringa = InputComTimeOut(5)
'    'Debug.Print "echo"; Stringa
'    Stringa = InputComTimeOut(5)
'    'Debug.Print Stringa
'    If InStr(Stringa, "?") Then
'        'Debug.Print "? Errore!"
'        MsgBox ("Errore GASCARD II")
'        Exit Sub
'    End If
'    If InStr(Stringa, "TimeOut") Then
'        'Debug.Print "? Errore! timeout"
'        MsgBox ("Lo spettrometro non risponde!")
'        Exit Sub
'    End If
'
'    'Debug.Print "Ready to Start"
    
    'AFGraphic1.Cls
    OnComm = True
    'fMain.mscomm1.RThreshold = 1
    
    'alternativa
    'ciclo
    StartTime = Timer
    MeasTime = 0
    CO2Index = 0
    Linea = ""
    Stringa = ""
    CO2hex = ""
    CO2 = 0
    'CO2MeasRaw = 0
    
        Do
            'Linea = InputComTimeOut(5)
            'Debug.Print Len(Linea)
            'Salta le linee incomplete
            Select Case CommType
            
                Case "GascardII"
                    Linea = InputComTimeOut(5)
                    If Len(Linea) < 41 Then
                        
                        If Linea = "0" Then
                            ' se la scheda non risponde manda un vbcr per risvegliarla!
                            fMain.MSComm1.Output = vbCr
                            Debug.Print "Risveglio!"
                        End If
                        GoTo NextLine
                    End If
                    'Prende i primi 4 caratteri che rappresentano la misura
                    'Dopo che la scheda è stata opportunamente settata prima.
                    Stringa = Left$(Linea, 4)
                    CO2hex = "&H" & 0 & Trim(Stringa) ' Mid$(Linea, CrLfIndex - 4, 4)
                    'Lo zero serve ad evitare che CSng si impalli per una stringa nulla
            
                    'Debug.Print CO2hex; vbTab;
                    'If CLng(Stringa) = 0 Then
                    '    CO2 = 0
                    'Else
                         CO2 = CSng(CO2hex)
                         CO2MeasRaw(CO2Index) = 0
                    'End If
                Case "Santino"
                    Linea = InputComTimeOutSantino(3)
                    If Linea <> "0" And Len(Linea) = 8 Then
                        Debug.Print "bStart Santino01 lenlinea="; Len(Linea)
                        FirstByte = Left(Linea, 1)
                        SecondByte = Mid(Linea, 2, 1)
                        'Debug.Print FirstByte; " "; SecondByte; " ";
                        MSB = Asc(FirstByte)
                        LSB = Asc(SecondByte)
                        'Debug.Print MSB; " "; LSB; " ";
                        Stringa = Left$(Linea, 2)
                        'Debug.Print Stringa; " ";
                        Stringa = SwapString(Stringa)
                        Word = bytes2long(Stringa)
                        'Debug.Print Word
'                        i = InStr(Linea, ",")
'                        Stringa = Left(Linea, i)
'                        Debug.Print Linea
                        CO2 = adc2value3(Word)
'                        CO2 = CO2 * 1 'fattore di conversione!!!!
                        CO2MeasRaw(CO2Index) = Word
                        
                        'Prendiamo il tempo
                        Stringa = Mid(Linea, 5, 2)
                        Stringa = SwapString(Stringa)
                        Word = bytes2long(Stringa)
                        
                        If SantinoStartTime = 0 Then
                            SantinoStartTime = Word
                        End If
                        SantinoTime = Word - SantinoStartTime
                        Debug.Print Word, SantinoTime
                    Else
                        'MSComm1.InBufferCount = 0
                        Debug.Print "Line zero "; Len(Linea)
                    End If
            End Select
            Debug.Print "bStart010 CO2="; CO2; " " '; vbTab;
            'Debug.Print CO2 & " ";
            CO2 = CO2 * FattoreScheda '/ 10000 '0.01 '/ 10000 * 100
            Debug.Print CO2
            
            CO2Meas(CO2Index) = CO2
            Select Case CommType
            
                Case "GascardII"
            
                    CO2MeasAr(CO2Index, 0) = MeasTime
                    MeasTime = MeasTime + Interval
                Case "Santino"
                    CO2MeasAr(CO2Index, 0) = SantinoTime
            End Select
            CO2MeasAr(CO2Index, 1) = CSng(CO2)
            'Debug.Print CO2MeasAr(CO2Index, 0), CO2MeasAr(CO2Index, 1)
            CO2Index = CO2Index + 1
            'DatiGrafico(iGrafico) = Int(CO2 * Scala)
            'Debug.Print DatiGrafico(iGrafico)
            'iGrafico = iGrafico + 1
            'AFGraphic1.SetPixel iGrafico, 100 - DatiGrafico(iGrafico - 1), 1
            'AFlTime.Caption = Int((Timer - StartTime) / 1000)
            lCoord.Caption = CO2
            MSChart1.ChartData = CO2MeasAr
            DoEvents
            'CheckGraph     'Questa routine dovrebbe visualizzare gli ultimi dati
NextLine:
        Loop Until OnComm = False
    
    Exit Sub
GestioneErrore:
    'Stringa = Err.Description + " in bStart_Click at" + Str(CO2Index) + " with " + Str(CO2) + "," + CO2hex + "," + Str(100 - DatiGrafico(iGrafico - 1))
    Stringa = Err.Description + " in bStart_Click at" + Str(CO2Index) + " with " + Str(CO2) + ",<" + CO2hex + ">," + Stringa + vbCrLf + Linea
    MsgBox Stringa

End Sub

Private Sub bStop_Click()
    If MeasStarted = False Then
        MsgBox ("Press Start first!")
        Exit Sub
    End If
    'AFTimer1.Interval = 0
    OnComm = False
'    CloseCom
'    fMain.mscomm1.RThreshold = 1
'    'fMain.mscomm1.SThreshold = 1
    MeasDone = True
    MeasSaved = False
    MeasStarted = False
    FirstClickOnGraph = False
    lCoord.Caption = "Stopped"
    CloseCom
End Sub

Private Sub bYmin_Click()
    Dim min As Integer
    On Error GoTo fine
    min = InputBox("Minimum Y")
    MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
    MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.Minimum = min
fine:
End Sub

Private Sub bYmax_Click()
    Dim max As Integer
    On Error GoTo fine
    max = InputBox("Maximum Y")
    MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
    MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = max
fine:
End Sub

Private Sub bYauto_Click()
    MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.Auto = True
End Sub

Private Sub bXmin_Click()
    Dim min As Integer
    On Error GoTo fine
    min = InputBox("Minimum X")
    MSChart1.Plot.Axis(VtChAxisIdX).ValueScale.Auto = False
    MSChart1.Plot.Axis(VtChAxisIdX).ValueScale.Minimum = min
fine:
End Sub

Private Sub bXmax_Click()
    Dim max As Integer
    On Error GoTo fine
    max = InputBox("Maximum X")
    MSChart1.Plot.Axis(VtChAxisIdX).ValueScale.Auto = False
    MSChart1.Plot.Axis(VtChAxisIdX).ValueScale.Maximum = max
fine:
End Sub

Private Sub bXauto_Click()
    MSChart1.Plot.Axis(VtChAxisIdX).ValueScale.Auto = True
End Sub

Private Sub MSChart1_PointSelected(Series As Integer, DataPoint As Integer, MouseFlags As Integer, Cancel As Integer)

    lCoord = Str(Series) + " " + Str(DataPoint)
    GetPoints DataPoint, DataPoint, CO2MeasAr(DataPoint, 0), CO2MeasAr(DataPoint, 1)
End Sub

