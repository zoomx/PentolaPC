VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
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
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   5760
      Width           =   735
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8520
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton bDownload 
      Caption         =   "&Download"
      Height          =   375
      Left            =   8400
      TabIndex        =   6
      Top             =   5760
      Width           =   975
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   7920
      Top             =   4800
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
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   5760
      Width           =   615
   End
   Begin VB.CommandButton bEnd 
      Caption         =   "&End"
      Height          =   375
      Left            =   9480
      TabIndex        =   3
      Top             =   5760
      Width           =   855
   End
   Begin VB.CommandButton bSave 
      Caption         =   "S&ave"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   5760
      Width           =   735
   End
   Begin VB.CommandButton bStop 
      Caption         =   "S&top"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   5760
      Width           =   615
   End
   Begin VB.CommandButton bStart 
      Caption         =   "&Start"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   5760
      Width           =   615
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
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   375
      Left            =   6480
      TabIndex        =   7
      Top             =   5760
      Width           =   1695
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    INIFile = App.Path + "\" + App.EXEName + ".ini"
    MSComm1.CommPort = 1
    MSComm1.Settings = "9600,n,8,1"
    MSComm1.InBufferSize = 16387
    'In win2000 è il massimo.
    'In win9x si può usare 32767
    ReDim CO2MeasAr(2000, 1)
End Sub

Private Sub bDownload_Click()
'    Me.Hide
'    Form1.Show
End Sub

Private Sub bEnd_Click()
    End
End Sub

Private Sub bSave_Click()
    Dim i As Integer

    If MeasDone = False Then Exit Sub
    CommonDialog1.ShowSave
    Open CommonDialog1.filename For Output As #1
    
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
    NomeFile = NomeFile + Stringa + Location



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
    CommonDialog1.ShowOpen
    Open CommonDialog1.filename For Input As #1
    Do
        Input #1, Linea
        Stringa = Left(Linea, 7)
    Loop Until Stringa = "Samples"
    Stringa = Mid(Linea, 10, Len(Linea) - 9)
    Samples = Val(Stringa)
    ReDim CO2MeasAr(Samples - 1, 1)
    Input #1, Linea
    On Error GoTo continua
    For i = 0 To Samples - 1 'UBound(CO2MeasAr)

        Input #1, CO2MeasAr(i, 0)
        'Print #1, ",";

        Input #1, CO2MeasAr(i, 1)


    Next
    On Error GoTo 0
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
    
    'On Error GoTo GestioneErrore
    If SetupDone = False Then
        MsgBox ("Press Setup first!")
        Exit Sub
    End If
    
        
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
        Case "Gascard II 10%"
            FattoreScheda = 10
            FondoScala = 10000
            Scala = 100 / FondoScala
        Case "Gascard II 3%"
            FattoreScheda = 3
            FondoScala = 30000
            Scala = 100 / FondoScala
        Case "Gascard II 1%"
            FattoreScheda = 1
            FondoScala = 10000
            Scala = 100 / FondoScala
    End Select
    
    'Scala = 0.01 '2000 / 100
    'AFTimer1.Interval = 1
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
    
    Stringa = InputComTimeOut(5)
    'Debug.Print "1 "; Stringa
    If Stringa = "TimeOut" Then
        'Debug.Print "timeout!"
        MSComm1.Output = vbCrLf
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
    If InStr(Stringa, "TimeOut") Then
        'Debug.Print "? Errore! timeout"
        MsgBox ("Lo spettrometro non risponde!")
        Exit Sub
    End If

    'Debug.Print "Ready to Start"
    'mscomm1.Output = vbCrLf
'    For x = 1 To 100
'        'AFGraphic1.SetPixel(x, x, 1)
'        AFGraphic1.SetPixel x, x, 1
'    Next
    
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
        MeasTime = MeasTime + 0.125
        CO2MeasAr(CO2Index, 1) = CSng(CO2)
        Debug.Print CO2MeasAr(CO2Index, 0), CO2MeasAr(CO2Index, 1)
        CO2Index = CO2Index + 1
        'DatiGrafico(iGrafico) = Int(CO2 * Scala)
        'Debug.Print DatiGrafico(iGrafico)
        'iGrafico = iGrafico + 1
        'AFGraphic1.SetPixel iGrafico, 100 - DatiGrafico(iGrafico - 1), 1
        'AFlTime.Caption = Int((Timer - StartTime) / 1000)
        lCoord.Caption = CO2
        MSChart1.ChartData = CO2MeasAr
        DoEvents
        'CheckGraph
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

End Sub

Private Sub bYmin_Click()
    Dim min As Integer
    min = InputBox("Minimum Y")
    MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
    MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.Minimum = min
End Sub

Private Sub bYmax_Click()
    Dim max As Integer
    max = InputBox("Maximum Y")
    MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
    MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = max
End Sub

Private Sub bYauto_Click()
    MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.Auto = True
End Sub

Private Sub bXmin_Click()
    Dim min As Integer
    min = InputBox("Minimum X")
    MSChart1.Plot.Axis(VtChAxisIdX).ValueScale.Auto = False
    MSChart1.Plot.Axis(VtChAxisIdX).ValueScale.Minimum = min
End Sub

Private Sub bXmax_Click()
    Dim max As Integer
    max = InputBox("Maximum X")
    MSChart1.Plot.Axis(VtChAxisIdX).ValueScale.Auto = False
    MSChart1.Plot.Axis(VtChAxisIdX).ValueScale.Maximum = max
End Sub

Private Sub bXauto_Click()
    MSChart1.Plot.Axis(VtChAxisIdX).ValueScale.Auto = True
End Sub

Private Sub MSChart1_PointSelected(Series As Integer, DataPoint As Integer, MouseFlags As Integer, Cancel As Integer)

    lCoord = Str(Series) + " " + Str(DataPoint)
    GetPoints DataPoint, DataPoint, CO2MeasAr(DataPoint, 0), CO2MeasAr(DataPoint, 1)
End Sub
