VERSION 5.00
Object = "{9A0F0269-9DD2-4928-9749-BB502308F61B}#1.0#0"; "GraphLite.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   10230
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton bOk 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   4920
      TabIndex        =   3
      Top             =   5400
      Width           =   615
   End
   Begin VB.CommandButton bStart 
      Caption         =   "&Start"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   5400
      Width           =   855
   End
   Begin VB.CommandButton bStop 
      Caption         =   "S&top"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   5400
      Width           =   855
   End
   Begin GraphLite98.GraphLite GraphLite1 
      Height          =   5000
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   8811
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   1
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bOk_Click()
    Me.Hide
    Unload Me
    fMain.Show

End Sub

Private Sub bStart_Click()
    Dim x As Integer
    Dim y As Integer
    
    Dim Linea As String         'Stringa ricevuta dalla RS232
    Dim CO2hex As String        'CO2 in hex
    Dim CO2 As Single           'CO2
    
    'On Error GoTo GestioneErrore
'    If SetupDone = False Then
'        MsgBox ("Press Setup first!")
'        Exit Sub
'    End If
    
        

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
    
    'MSChart1.Repaint = False
    'MSChart1.RowCount = 2
    GraphLite1.chartType = VtChChartType2dXY
    'MSChart1.chartType = VtChChartType2dXY
    
    'MSChart1.Plot.UniformAxis = False
    'MSChart1.ChartData = Dati
    'MSChart1.Repaint = True
    GraphLite1
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
        MeasTime = MeasTime + 0.25
        CO2MeasAr(CO2Index, 1) = CSng(CO2)
        Debug.Print CO2MeasAr(CO2Index, 0), CO2MeasAr(CO2Index, 2)
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

Private Sub GraphLite1_Click()
    GraphLite1.CurrentX
End Sub
