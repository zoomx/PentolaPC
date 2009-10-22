Attribute VB_Name = "VarGlobali"
Option Explicit

Type pdbRecord
    NomeFile As String
    data As String
End Type

'Variabili globali per il programma

Public Intero As Integer
Public Lungo As Long
Public Stringa As String

Public ComPort As Long
Public OnComm As Boolean
Public Messaggio As String
'Public Intero As Integer
'Public Lungo As Long
'Public Float As Single
'Public Dfloat As Double
'Public Stringa As String
Public ComOk As Boolean
Public Collegato As Boolean
Public FileOut As String
Public filename As String
Public Const Vero As Boolean = True
Public Const Falso As Boolean = False
Public CO2MeasRaw(2000) As Long
Public CO2Meas(2000) As Single
'Public CO2MeasAr(2000, 1) As Single
Public CO2MeasAr() As Single
Public CO2Index As Integer
Public CO2IndexAr(2) As Integer
Public MeasTime As Single
Public DatiGrafico(160) As Integer
Public iGrafico As Integer
Public Scala As Single
Public FondoScala As Single
Public FattoreScheda As Single
Public dbHandle As Long
Public myRecord As pdbRecord
Public StartTime As Single
Public StopTime As Single
Public TimeNow As Single

Public NomeFile As String
Public Location As String
Public AtmPressure As String
Public AtmTemp As String
Public WindVelocity As String
Public CO2sensor As String
Public LocData As Boolean
Public MeasDone As Boolean
Public SetupDone As Boolean
Public MeasSaved As Boolean
Public MeasStarted As Boolean
Public FirstClickOnGraph As Boolean
Public GettingPoint As String
Public LongMeasure As Boolean   'For measures that take one hour
Public Dummy As String

Public x1 As Integer
Public x2 As Integer
Public y1 As Integer
Public y2 As Integer

Public CommType As String
