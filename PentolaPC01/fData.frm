VERSION 5.00
Begin VB.Form fData 
   Caption         =   "Location data"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "fData.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox tComPort 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2160
      TabIndex        =   16
      Text            =   "1"
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox tKchamber 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2160
      TabIndex        =   13
      Text            =   "14"
      Top             =   1920
      Width           =   1455
   End
   Begin VB.ComboBox ComboSensors 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   2160
      TabIndex        =   12
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox tWind 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2160
      TabIndex        =   11
      Text            =   "2"
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox tHum 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2160
      TabIndex        =   10
      Text            =   "50"
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox tTemper 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2160
      TabIndex        =   9
      Text            =   "18"
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox tPressure 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2160
      TabIndex        =   8
      Text            =   "1013"
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox tLocation 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2160
      TabIndex        =   7
      Text            =   "Palermo"
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton bOk 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   3840
      TabIndex        =   6
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label8 
      Caption         =   "ComPort"
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "K Acc. Chamber"
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "CO2 sensor"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Wind velocity"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Air Humidity"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Atmosp. Temperature"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Atmosp. Pressure"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Location"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "fData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bOk_Click()
    If CO2sensor = "" Then
        MsgBox ("Select CO2 sensor!")
        Exit Sub
    End If
    Location = tLocation.Text
    AtmPressure = tPressure.Text
    AtmTemp = tTemper.Text
    WindVelocity = tWind.Text
    ComPort = Val(tComPort.Text)
    Kchamber = tKchamber.Text
    CloseCom
    fMain.MSComm1.CommPort = ComPort
    WriteINI "CommSettings", "Comport", ComPort, INIFile
    WriteINI "Sensor", "CardType", CO2sensor, INIFile
    WriteINI "Sensor", "K", Kchamber, INIFile
    
    SetupDone = True
    'qui scrivo che scheda uso nella label del mai
    fMain.lCardType.Caption = CO2sensor + " K=" + Str(Kchamber)
    Me.Hide
    Unload Me
    fMain.Show

End Sub

Private Sub ComboSensors_Click()
    CO2sensor = ComboSensors.Text
    fMain.lGascard.Caption = CO2sensor
End Sub

Private Sub Form_Load()
    ComboSensors.AddItem "Gascard II 100%"
    ComboSensors.AddItem "Gascard II 30%"
    ComboSensors.AddItem "Gascard II 10%"
    ComboSensors.AddItem "Gascard II 5%"
    ComboSensors.AddItem "Gascard II 3%"
    ComboSensors.AddItem "Gascard II 1%"
    ComboSensors.AddItem "Gascard II 3000 ppm"
    'ComboSensors.AddItem "Licor 820"
    'ComboSensors.AddItem "Mastrolia 10000"
    'ComboSensors.AddItem "Mastrolia 100%"
    ComboSensors.AddItem "Santino 10000ppm"
    ComboSensors.AddItem "Santino 100%"
    
    If ComPort > 0 Then tComPort.Text = ComPort
    If CO2sensor <> "" Then ComboSensors.SelText = CO2sensor
    If Kchamber <> 0 Then tKchamber.Text = Kchamber
 
    
End Sub
