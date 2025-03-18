Sub ScriptIW39()

Dim SapGuiAuto
Dim Applicationx
Dim Connection
Dim session
Dim isSessionBusy As Boolean


'Definir pasta downloads para qualquer usuário
PastaDownload = Environ("USERPROFILE") & "\Downloads"

If Not IsObject(Applicationx) Then
   Set SapGuiAuto = GetObject("SAPGUI")
   Set Applicationx = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(Connection) Then
   Set Connection = Applicationx.Children(0)
End If
If Not IsObject(session) Then
   Set session = Connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session, "on"
   WScript.ConnectObject Applicationx, "on"
End If

'Acessar IW39'
 session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nIW39"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtDATUV").Text = "01.01.2024"
session.findById("wnd[0]/usr/ctxtDATUB").Text = ""
session.findById("wnd[0]/usr/ctxtIWERK-LOW").Text = "2298"
session.findById("wnd[0]/usr/ctxtVARIANT").Text = "/indiciw38"
session.findById("wnd[0]/usr/ctxtVARIANT").SetFocus
session.findById("wnd[0]/usr/ctxtVARIANT").caretPosition = 10
session.findById("wnd[0]/tbar[1]/btn[8]").press

'Salvar Excel .xls'
session.findById("wnd[0]/mbar/menu[0]/menu[11]/menu[2]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = PastaDownload
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "2298 - IW39.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 12
session.findById("wnd[1]/tbar[0]/btn[11]").press


 Dim wb As Workbook
    Dim Arquivo
    
    ' Caminho do arquivo .xls
    Arquivo = PastaDownload & "\2298 - IW39.xls"
   
    ' Abrir o arquivo .xls
    Set wb = Workbooks.Open(Arquivo)
    
  
  'Transforma dados em tabela
    Rows("1:3").Select
    Selection.Delete Shift:=xlUp
    Rows("2:2").Select
    Selection.Delete Shift:=xlUp
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
            
  'Salvar como .xlsx no Sharepoint
  ActiveWorkbook.SaveAs Filename:= _
    "https://suzano.sharepoint.com/sites/ProjetoCerrado-ConfiabilidadeInovao/Documentos%20Compartilhados/99%20-%20Indicadores%20Industriais/01%20-%20Indicadores%20Ribas/02%20-%20Painel%20Extra%C3%A7%C3%A3o%20SAP/2298%20-%20IW39.xlsx" _
        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
                   
    'Fechar o arquivo .xls
    wb.Close SaveChanges:=False

    ' Limpar a variável
    Set wb = Nothing

End Sub
