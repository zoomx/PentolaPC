Attribute VB_Name = "SpectrometersFuncions"
Option Explicit

Public Function InitCards() As Boolean
    InitCards = False
    Select Case CommType
        Case "GascardII"
            'fMain.MSComm1.Output = vbCr
            fMain.MSComm1.Output = "PT000"
            fMain.MSComm1.InBufferCount = 0
            fMain.MSComm1.Output = "E00"
            Stringa = InputComTimeOut(5)
            If InStr(Stringa, "?") Then
                'Debug.Print "? Errore!"
                MsgBox ("Errore GASCARD II")
                InitCards = False
                Exit Function
            End If
            If InStr(Stringa, "TimeOut") Then
                'Debug.Print "? Errore! timeout"
                MsgBox ("Lo spettrometro non risponde!")
                InitCards = False
                Exit Function
            End If

'    'Debug.Print "Ready to Start"

            If Stringa = "0" Then
                'Timeout
                InitCards = False
            Else
                InitCards = True
            End If

        Case "Mastrolia"
            Stringa = InputComTimeOut(5)
            If Stringa = "0" Then
                'Timeout
                InitCards = False
            Else
                InitCards = True
            End If

        Case "Santino"
            'non fa niente per adesso
            
    End Select
    
    Debug.Print "Init card done!"
End Function
