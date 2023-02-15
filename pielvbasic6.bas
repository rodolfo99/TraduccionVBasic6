

Public Function mostrarNombreControles(Mi As Form)

    Dim obj, dbs, controlForm As Object
                
'Open mi.Name For Append As #1
        
        Dim rsDic As ADODB.Recordset
        Dim strdriver As String
140        strdriver = "PostgreSQL ODBC Driver(UNICODE)"
150        cadenaConexionGlobal = "DRIVER={" & strdriver & "};SERVER=148.213.21.76;PORT=9002" & _
                                  ";DATABASE=base9;UID=postgres;PWD={latinmix54};ByteaAsLongVarBinary=1;ApplicationName=Aplicacion;"
        
        
        For Each controlForm In Mi.Controls
          If (TypeOf controlForm Is Label) Or _
               (TypeOf controlForm Is Frame) Or _
               (TypeOf controlForm Is CheckBox) Or _
               (TypeOf controlForm Is OptionButton) Or _
               (TypeOf controlForm Is Menu) Or _
               (TypeOf controlForm Is CommandButton) Or _
               (TypeOf controlForm Is MyButton) Or _
               (TypeOf controlForm Is MyFrame) Then
           ' MsgBox controlForm.Name
            'Print #1, mi.Name & "," & controlForm.Name & "," & controlForm.Caption
         
            
            Set rsDic = ExecuteSQLrs("SELECT * FROM config_siabuc.diccionariodatos WHERE formulario='" & Mi.Name & "' AND control='" & controlForm.Name & "' AND texto like '%" & Replace(controlForm.Caption, "'", "''") & "%' ORDER BY id_dic ASC")
            On Error Resume Next
            controlForm.Caption = rsDic.fields("sustituto").Value
            
            rsDic.Close
            
          End If
        Next

'Close #1
End Function

Public Function guardarNombreControles(Mi As Form)

    Dim controlForm As Object
    Dim strSQL As String
    Dim errorLocal As String
    Dim rsDic As ADODB.Recordset
'Open "c:\" & Mi.Name For Append As #1
      Dim strdriver As String
140        strdriver = "PostgreSQL ODBC Driver(UNICODE)"
150        cadenaConexionGlobal = "DRIVER={" & strdriver & "};SERVER=127.0.0.1;PORT=9002" & _
                                  ";DATABASE=base9;UID=postgres;PWD={xxx};ByteaAsLongVarBinary=1;ApplicationName=Aplicacion;"
          
        
   
   Set rsDic = ExecuteSQLrs("SELECT * FROM config_siabuc.diccionariodatos WHERE formulario='" & Mi.Name & "'")
   On Error Resume Next
   If rsDic Is Nothing Then
        rsDic.Close
        For Each controlForm In Mi.Controls
          If (TypeOf controlForm Is Label) Or _
               (TypeOf controlForm Is Frame) Or _
               (TypeOf controlForm Is CheckBox) Or _
               (TypeOf controlForm Is OptionButton) Or _
               (TypeOf controlForm Is Menu) Or _
               (TypeOf controlForm Is CommandButton) Or _
               (TypeOf controlForm Is MyButton) Or _
               (TypeOf controlForm Is MyFrame) Then
           ' MsgBox controlForm.Name
            'Print #1, Mi.Name & "," & controlForm.Name & "," & controlForm.Caption
            strSQL = "INSERT INTO config_siabuc.diccionariodatos(modulo,formulario, control, texto,  sustituto) VALUES ('" & nomModulo & "','" & Mi.Name & "', '" & controlForm.Name & "', '" & controlForm.Caption & "', '@" & controlForm.Caption & "' ) ;"
                Call ExecuteSQLaction(strSQL, , errorLocal)
                  
                  
          End If
        Next
    End If
'Close #1
               
                
End Function

