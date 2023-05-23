Attribute VB_Name = "SAP_TX"
Dim Session As SAPFEWSELib.GuiSession
Dim StatusBar As SAPFEWSELib.GuiStatusbar

Sub ejecutar_TX(condicion As String, clientes As Range, fecha_desde As String, archivo_path As String, archivo_nombre As String)
    
    ' obtengo la sesion de SAP o inicio
    Set Session = SAP_session
    Set StatusBar = Session.FindById("wnd[0]/sbar")
    
    Session.StartTransaction ("TX_SAP")
    ' comandos de TX aquí
    
    Session.EndTransaction
    
End Sub
