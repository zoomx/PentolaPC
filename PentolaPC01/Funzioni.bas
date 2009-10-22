Attribute VB_Name = "Funzioni"
Option Explicit

Public Function InputComTimeOut(TimeOut As Integer) As String
'Attende un input il cui terminatore e' vbLF
'Con TIMEOUT
    Dim TimeStop As Long
    Dim Linea As String
    Dim Dummy As String
        
        'set ishell=
        'TimeOut = TimeOut * 1000 'passiamo ai millisecondi
        
        TimeStop = Timer + TimeOut
        'Debug.Print iShelll.GetTimeMS
        fMain.MSComm1.InputLen = 1
        Do
            DoEvents

        Loop Until (fMain.MSComm1.InBufferCount >= 1) Or (Timer > TimeStop)
        If fMain.MSComm1.InBufferCount >= 1 Then
            Linea = ""
            Dummy = ""
            TimeStop = Timer + TimeOut ' Imposta l'ora di fine
            Do Until Dummy = vbLf Or (Timer > TimeStop)
                DoEvents
                Dummy = fMain.MSComm1.Input
                Linea = Linea + Dummy
            Loop
        Else
            Linea = "0"
        End If
        InputComTimeOut = Linea

End Function

Public Sub OpenCom()
    'Apre la porta com
    'Se e' andata bene ComOk e' True altrimenti e' False
    Dim Msg As String

    On Error GoTo ErroreCom
    ComOk = False
    'Apre la porta seriale se non è già aperta
    If fMain.MSComm1.PortOpen = False Then fMain.MSComm1.PortOpen = True
    ComOk = True
    Exit Sub
ErroreCom:
    Select Case Err.Number
        Case 8005  'La Com è già aperta
            Msg = "Errore la porta Com" + Str$(ComPort) + " è già in uso"
            MsgBox Msg, vbOKOnly, "Errore"

            
            Err.Clear   ' Cancella i campi dell'oggetto
            ComOk = False
            Exit Sub
        Case 8002
            Msg = "Errore la porta Com" + Str$(ComPort) + " non esiste!"
            MsgBox Msg, vbOKOnly, "Errore"

            Err.Clear   ' Cancella i campi dell'oggetto
            ComOk = False
            Exit Sub
        Case Else
            Msg = Err.Description
            MsgBox Msg, vbOKOnly, "Errore"


            Exit Sub
    End Select

End Sub

Public Sub CloseCom()
    'Chiude la porta seriale se non è già chiusa
    fMain.MSComm1.InBufferCount = 0
    If fMain.MSComm1.PortOpen = True Then fMain.MSComm1.PortOpen = False
End Sub
Public Function Val2(Valore As String) As Single
'Simile alla val ma per separatore decimale usa sia il
'punto che la virgola
    Dim ip As Integer
    Dim iv As Integer
    Dim lStringa As Integer
    Dim temp As Single
    Dim Stringa As String
    
    Stringa = CStr(Valore)
    'C'è il punto?
    ip = InStr(Stringa, ".")
    'C'è la virgola?
    iv = InStr(Stringa, ",")
    lStringa = Len(Stringa)
    If iv <> 0 Then 'Se c'è la virgola la sostituisce col punto
        Stringa = Left(Stringa, iv - 1) + "." + Right(Stringa, lStringa - iv)
        ip = iv
    End If
    temp = CSng(Stringa)
    'If ip <> 0 And iv <> 0 Then
    'Se ci sono tutte e due?
    Val2 = temp
End Function

Public Sub CheckGraph()
    Dim i As Integer
    On Error GoTo GestioneErrore
    If iGrafico >= 160 Then
        'fMain.AFGraphic1.Cls
        
'        For i = 2 To 160
'            DatiGrafico(i - 1) = DatiGrafico(i)
'            fMain.AFGraphic1.SetPixel i - 1, 100 - DatiGrafico(i - 1), 1
'        Next i
        For i = 51 To 160
            DatiGrafico(i - 50) = DatiGrafico(i)
            'fMain.AFGraphic1.SetPixel i - 50, 100 - DatiGrafico(i - 50), 1
        Next i

        iGrafico = 110
    End If
    
    Exit Sub
GestioneErrore:
    Stringa = Err.Description + " in CheckGraph"
    MsgBox Stringa

End Sub

Public Function CreateDatabase() As Long
    Dim dbHandle As Long
'    #If APPFORGE Then
'        'Create new database (if on the device)
'        'dbHandle = PDBCreateDatabase("flowmeas", lType, lCreator)
'        dbHandle = PDBCreateDatabase("flowmeas", _
'        PalmIDtoLong("DATA"), PalmIDtoLong("INGV"))
'    #Else
'        'Create new database (if on the PC)
'        dbHandle = PDBCreateDatabase(App.Path & "\flowmeas", _
'        PalmIDtoLong("DATA"), PalmIDtoLong("INGV"))
'    #End If
'        'Create the table (db as Long, TableName as String,
'        'FieldString As String) as Long
'        PDBCreateTable dbHandle, "flowmeas", "Filename String, Data String"
        CreateDatabase = dbHandle
End Function

Private Function PalmIDtoLong(PalmID As String) As Long
    Dim myLng As Long, Counter As Integer
    If Len(PalmID) = 4 Then
        For Counter = 1 To Len(PalmID)
            myLng = myLng * 256 + Asc(Mid(PalmID, Counter, 1))
        Next Counter
        PalmIDtoLong = myLng
    Else
        PalmIDtoLong = 0
    End If
End Function

Public Function OpenDatabase() As Boolean
    Dim dbHandle As Long
'    lType = PalmIDtoLong("DATA")
'    lCreator = PalmIDtoLong("INGV")

        ' Open the database
'        #If APPFORGE Then
'            dbHandle = PDBOpen(Byfilename, "flowmeas", 0, 0, 0, 0, afModeReadWrite)
'        #Else
'            dbHandle = PDBOpen(Byfilename, App.Path & "\flowmeas", 0, 0, 0, 0, afModeReadWrite)
'        #End If
'
'        If dbHandle <> 0 Then
'                'We successfully opened the database
'                OpenDatabase = True
'        Else
'                'We failed to open the database
'                'MsgBox "No database found. Creating new database.", vbOKOnly
'                'Call CreateAnimalsDatabase
'
'
'                'OpenAnimalsDatabase = True
'
'        End If
        

End Function

Public Sub GetPoints(x As Integer, y As Integer, Time As Single, value As Single)
    
    Static index1 As Double
    'Static Time1 As Double
    Static x1 As Integer
    Static value1 As Double
    Dim index2 As Double
    
    Dim x2 As Integer
    Dim value2 As Double
    Dim m As Double
    Dim m2 As Double
    Dim q As Double
    Dim q2 As Double
    'Dim temp As Single
    Dim i As Integer
    Dim SommaX As Double
    Dim SommaX2 As Double
    Dim SommaY As Double
    Dim SommaY2 As Double
    Dim SommaXY As Double
    Dim a As Double
    Dim r2 As Double
    Dim n As Integer
    Dim SommaXSommaYn As Double
    Dim SommaX2n As Double
    Dim SommaY2n As Double
    Dim SommaMista As Double
    Dim SommaXX As Double
    Dim SommaYY As Double
    
    Dim NP As Double
    Dim a3 As Double
    Dim b3 As Double
    Dim r3 As Double
    
    If GettingPoint = "first" Then
        x1 = x
'        y1 = y
        index1 = Time
        value1 = value
        GettingPoint = "second"
        fMain.Label1.Caption = "First point selected " + Str(index1) + " " + Str(value)
    Else
        x2 = x
'        y2 = y
        index2 = Time
        value2 = value
        GettingPoint = "Done"
        fMain.Label1.Caption = "Done"
    End If
    Debug.Print GettingPoint
    If GettingPoint = "Done" Then
    'Se sono stati presi entrambi i punti
    'calcola la pendenza della retta
'        If x2 - x1 = 0 Then
'            m = 0 'sarebbe infinito
'        Else
'            m = -(y2 - y1) / (x2 - x1)
'        End If
        If index2 - index1 = 0 Then
            m2 = 0 'Anche qui sarebbe infinito
        Else
            m2 = (value2 - value1) / (index2 - index1)
        End If
        
'        q = y1 - m * x1
        q2 = value1 - m2 * index1
        'Debug.Print "retta -> q="; q; " m="; m
        'Metodo WEST
        SommaX = 0
        SommaX2 = 0
        SommaY = 0
        SommaY2 = 0
        SommaXY = 0
        n = x2 - x1
        If n = 0 Then
            GettingPoint = "first"
            fMain.Label1.Caption = "Select first point"
            Exit Sub
        End If
        For i = x1 To x2
            SommaX = SommaX + CO2MeasAr(i, 0)
            SommaX2 = SommaX2 + CO2MeasAr(i, 0) * CO2MeasAr(i, 0)
            SommaY = SommaY + CO2MeasAr(i, 1)
            SommaY2 = SommaY2 + CO2MeasAr(i, 1) * CO2MeasAr(i, 1)
            SommaXY = SommaXY + CO2MeasAr(i, 0) * CO2MeasAr(i, 1)

        Next
        'Dedotto dalle formule
'        SommaXSommaYn = (SommaX * SommaY) / n
'        Debug.Print "SommaXSommaYn "; SommaXSommaYn
'        SommaX2n = (SommaX * SommaX) / n
'        Debug.Print "SommaX2n "; SommaX2n
'        SommaY2n = (SommaY * SommaY) / n
'        Debug.Print "SommaY2n "; SommaY2n
'        SommaMista = SommaXY - SommaXSommaYn
'        Debug.Print "SommaMista "; SommaMista
'        SommaXX = SommaX2 - SommaX2n
'        SommaYY = SommaY2 - SommaY2n
'        Debug.Print SommaMista, SommaXX, SommaYY
'        a = SommaMista / SommaXX
'        r2 = SommaMista * SommaMista / SommaXX * SommaYY
        
        'Copiato dal programma WEST
        
        SommaX2n = SommaX / n
        SommaY2n = SommaY / n
        If (SommaX2 - (SommaX ^ 2) / n) <> 0 Then
            a = (SommaXY - (SommaX * SommaY) / n) / (SommaX2 - (SommaX ^ 2) / n)
            If (SommaX2 - (SommaX ^ 2) / n) * (SommaY2 - (SommaY ^ 2) / n) <> 0 Then
                r2 = ((SommaXY - SommaX * SommaY / n) ^ 2) / ((SommaX2 - (SommaX ^ 2) / n) * (SommaY2 - (SommaY ^ 2) / n))
            Else
                r2 = 0
            End If
        Else
            a = 0
            r2 = 0
        End If
        If r2 < 0 Then
            r2 = 0
        End If
        If r2 > 1 Then
            r2 = 1
        End If

        'Copiato da Regressioni2
'        For i = N0 To N1
'            sommax = sommax + x(i)
'            sommay = sommay + y(i)
'            sommax2 = sommax2 + x(i) * x(i)
'            sommaxy = sommaxy + x(i) * y(i)
'            sommay2 = sommay2 + y(i) * y(i)
'        Next i
'
    NP = CSng(x2 - x1) + 1
    If (NP * SommaX2 - SommaX * SommaX) = 0 Then
            GettingPoint = "first"
            fMain.Label1.Caption = "Select first point"
            Exit Sub

    End If
    a3 = (NP * SommaXY - SommaX * SommaY) / (NP * SommaX2 - SommaX * SommaX)
    b3 = (SommaX2 * SommaY - SommaX * SommaXY) / (NP * SommaX2 - SommaX * SommaX)
    r3 = (NP * SommaXY - SommaX * SommaY) / Sqr((NP * SommaX2 - SommaX * SommaX) * (NP * SommaY2 - SommaY * SommaY))

        
        'fMain.lCoord.Caption = Str(m)
        
'        Stringa = "m con punti" + Str(m) + vbCrLf
'        Stringa = Stringa + "m con valori=" + Str(m2) + vbCrLf
        Stringa = "m con valori estremi=" + Str(m2) + vbCrLf
        Stringa = Stringa + "CO2=" + Str(m2 * 14) + vbCrLf
'        Stringa = Stringa + "metodo WEST" + vbCrLf
'        Stringa = Stringa + "a=" + Str(a) + vbCrLf
'        Stringa = Stringa + "R2=" + Str(r2) + vbCrLf
        Stringa = Stringa + "Metodo definitivo" + vbCrLf
        Stringa = Stringa + "a=" + Str(a3) + vbCrLf
        Stringa = Stringa + "R2=" + Str(r3) + vbCrLf
        Stringa = Stringa + "CO2=" + Str(a3 * 14)
        MsgBox Stringa
        GettingPoint = "first"
        fMain.Label1.Caption = ""
    End If
    
End Sub

Public Sub WaitSeconds(Seconds As Long)
    Dim Stime As Long
    Stime = Timer
    Do
        DoEvents
    Loop Until Timer - Stime > Seconds
End Sub

Public Function GetNameFromDir(Dir As String) As String
    Dim i As Long
    Dim lasti As Long
    Dim Dirr As String
    Dirr = Dir
    Do
        lasti = i
        i = InStr(Dir, "\")
        Dir = Right(Dir, Len(Dir) - i)
    Loop Until i = 0
    GetNameFromDir = Dir
End Function

Function sGetAppPath() As String
'*Returns the application path with a trailing \.      *
'*To use, call the function [SomeString=sGetAppPath()] *
Dim sTemp As String
        sTemp = App.Path
        If Right$(sTemp, 1) <> "\" Then sTemp = sTemp + "\"
        sGetAppPath = sTemp
End Function

Public Sub NewPath(Stringa As String)
'Cambia drive e path contemporaneamente
'Modificare per i drive di rete
'Es. NewPath "d:\temp"
    ChDrive (Left(Stringa, 3))
    ChDir (Stringa)
End Sub
Public Function adc2value(Valore_ADC As Long, Bitmin As Long, _
Bitmax As Long, valMax As Double, valMin As Double, valOff _
As Double) As Double
'From ADCount to Value

    Dim Valore As Double
    Valore = (Valore_ADC - Bitmin) / (Bitmax - Bitmin) * _
    (valMax - valMin) + valMin + valOff
    adc2value = Valore
    'Debug.Print "adc2value-->"; Valore
End Function


