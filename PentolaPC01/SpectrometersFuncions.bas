Attribute VB_Name = "SpectrometersFuncions"
Option Explicit

Public Function InitCards() As Boolean
    InitCards = False
    Select Case CommType
        Case "GascardII"
            fMain.MSComm1.Output = vbCrLf
            fMain.MSComm1.Output = "PT000"
            fMain.MSComm1.InBufferCount = 0
            fMain.MSComm1.Output = "E00"
            Stringa = InputComTimeOut(5)
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

    
    End Select
    

End Function
