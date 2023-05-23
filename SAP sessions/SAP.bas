Attribute VB_Name = "SAP"
Dim SapGui As Object
Dim SapGuiAPP As SAPFEWSELib.GuiApplication
Dim Connection As SAPFEWSELib.GuiConnection, ConnectionAux As SAPFEWSELib.GuiConnection
Dim Session As SAPFEWSELib.GuiSession, SessionAux As SAPFEWSELib.GuiSession

' Defino constantes
Const sapConnection As String = "SAP Test"
Const sapDir As String = "C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
Const saplogon As String = "saplogon.exe"

' defino las sesiones maximas que puedo usar
Const maxSessions As Integer = 3

Public Function SAP_session() As SAPFEWSELib.GuiSession
    
    Dim i As Integer
    If OpenSAPFrontEnd <> True Then
        MsgBox "Ocurrió un error al intentar iniciar SAP Logon", vbCritical, "Mensaje al Usuario"
    End If
    
    Do While (SapGuiAPP Is Nothing) Or (SapGui Is Nothing)
        Application.Wait (Now + TimeValue("00:00:01"))
        Set SapGui = GetObject("SAPGUI")
        Set SapGuiAPP = SapGui.GetScriptingEngine
    Loop
    
    If SapGuiAPP.Connections.Count > 0 Then
        For Each ConnectionAux In SapGuiAPP.Connections
            If ConnectionAux.Description = sapConnection Then
                Set Connection = SapGuiAPP.FindById(ConnectionAux.Name)
                Set ConnectionAux = Nothing
                Set Session = Connection.Sessions.Item(0)
                Exit For
            End If
        Next
    End If
    
    If Connection Is Nothing Then
        Set Connection = SapGuiAPP.OpenConnection(sapConnection)
        Application.Wait (Now + TimeValue("00:00:02"))
        Set Session = Connection.Sessions(0)
        GoTo conexionDeterminada
    End If
    
buscarSesiones:
    For Each SessionAux In Connection.Sessions
        If SessionAux.Busy Then
            GoTo ocupada
        ElseIf SessionAux.Info.Transaction = "SESSION_MANAGER" Then
            Set Session = Connection.FindById(SessionAux.Name)
            Set SessionAux = Nothing
            GoTo conexionDeterminada
        End If
ocupada:
    Next
    
    If checkMaxSessions(Connection) Then
        Session.CreateSession
        Application.Wait (Now + TimeValue("00:00:02"))
        GoTo buscarSesiones
    Else
        For i = 1 To Connection.Sessions.Count
            Set Session = Connection.Sessions(Connection.Sessions.Count - i)
            If Session.Busy = False Then Exit For
        Next i
    End If
    
    Set SessionAux = Nothing
    
conexionDeterminada:
    Set Connection = Nothing
    Set SapGuiAuto = Nothing
    Set SapGuiAPP = Nothing
        
    Set SAP_session = Session

End Function

Private Function OpenSAPFrontEnd() As Boolean

    On Error GoTo fallo
    If isOpenSAPFrontEnd() <> True Then
        vlArchivo = Shell(sapDir)
        Do
            Application.Wait (Now + TimeValue("00:00:01"))
        Loop Until isOpenSAPFrontEnd()
    End If
    OpenSAPFrontEnd = True
    Exit Function

fallo:
    OpenSAPFrontEnd = Err.Description
    OpenSAPFrontEnd = False
    
End Function

Private Function isOpenSAPFrontEnd() As Boolean

    target_process = Chr(34) & saplogon & Chr(34)
    
    Dim cObj As Object
    Dim proceso As Object
    
    On Error GoTo fallo
    Set cObj = GetObject("winmgmts://.")
    Set proceso = cObj.ExecQuery("SELECT * FROM Win32_Process WHERE Name = " & target_process)
    
    If proceso.Count > 0 Then
        isOpenSAPFrontEnd = True
        GoTo cerrar
    End If
    
fallo:
    isOpenSAPFrontEnd = False
cerrar:
    Set proceso = Nothing
    Set cObj = Nothing
    
End Function

Private Function checkMaxSessions(ByVal Connection As SAPFEWSELib.GuiConnection) As Boolean
    checkMaxSessions = IIf(Connection.Sessions.Count < maxSessions, True, False)
End Function
