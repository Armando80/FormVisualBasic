Dim Bandera As Boolean

Private Sub CommandButton1_Click()
    Me.TextBox1.Value = ""
    Me.TextBox2.Value = ""
    Me.TextBox3.Value = ""
    Me.TextBox4.Value = ""
    Me.TextBox5.Value = ""
    Me.TextBox6.Value = ""
    Me.TextBox7.Value = ""
    Me.TextBox8.Value = ""
    Me.TextBox9.Value = ""
    Me.TextBox10.Value = ""
    Me.TextBox11.Value = ""
    Me.TextBox12.Value = ""
    Me.TextBox13.Value = ""
    Me.TextBox14.Value = ""
End Sub

'bOTOn para insertar un reguistro
Private Sub CommandButton2_Click()
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("Data")
    Dim last_Row As Long
    last_Row = Application.WorksheetFunction.CountA(sh.Range("A:A"))
    'Validations-----------------------------------------------------
    If Me.TextBox1.Value = "" Then
    MsgBox "Por favor introduce la Fecha de Recibo", vbCritical
    Exit Sub
    End If
    If Me.TextBox2.Value = "" Then
    MsgBox "Por favor introduce el Número de Lote", vbCritical
    Exit Sub
    End If
    If Me.TextBox3.Value = "" Then
    MsgBox "Por favor introduce el Código SAP del rollo", vbCritical
    Exit Sub
    End If
    If Me.TextBox4.Value = "" Then
    MsgBox "Por favor introduce el motivo del rechazo", vbCritical
    Exit Sub
    End If
    If Me.TextBox5.Value = "" Then
    MsgBox "Por favor introduce la clave HC", vbCritical
    Exit Sub
    End If
    If Me.TextBox6.Value = "" Then
    MsgBox "Por favor introduce la Descripcion del HC", vbCritical
    Exit Sub
    End If
    If Me.TextBox7.Value = "" Then
    MsgBox "Por favor introduce el número de pedido", vbCritical
    Exit Sub
    End If
    If Me.TextBox8.Value = "" Then
    MsgBox "Por favor introduce la Fecha de Rechazo", vbCritical
    Exit Sub
    End If
    If Me.TextBox9.Value = "" Then
    MsgBox "Por favor introduce la Cantidad Rechazada", vbCritical
    Exit Sub
    End If
    If Me.TextBox10.Value = "" Then
    MsgBox "Por favor introduce el nombre del Proveedor del Metal", vbCritical
    Exit Sub
    End If
    If Me.TextBox11.Value = "" Then
    MsgBox "Por favor introduce descripción del código SAP", vbCritical
    Exit Sub
    End If
    If Me.TextBox12.Value = "" Then
    MsgBox "Por favor introduce el número de Factura/Remisión", vbCritical
    Exit Sub
    End If
    If Me.TextBox12.Value = "" Then
    MsgBox "Por favor introduce tú nombre!", vbCritical
    Exit Sub
    End If
    '-----------------------------------------------------------------
    sh.Range("A" & last_Row + 1).Value = "=Row()-1"
    sh.Range("B" & last_Row + 1).Value = Me.TextBox8.Value
    sh.Range("C" & last_Row + 1).Value = Me.TextBox1.Value
    sh.Range("D" & last_Row + 1).Value = Me.TextBox3.Value
    sh.Range("E" & last_Row + 1).Value = Me.TextBox11.Value
    sh.Range("F" & last_Row + 1).Value = Me.TextBox2.Value
    sh.Range("G" & last_Row + 1).Value = Me.TextBox9.Value
    sh.Range("H" & last_Row + 1).Value = Me.TextBox4.Value
    sh.Range("I" & last_Row + 1).Value = Me.TextBox10.Value
    sh.Range("J" & last_Row + 1).Value = Me.TextBox7.Value
    sh.Range("K" & last_Row + 1).Value = Me.TextBox12.Value
    sh.Range("L" & last_Row + 1).Value = Me.TextBox5.Value
    sh.Range("M" & last_Row + 1).Value = Me.TextBox6.Value
    sh.Range("N" & last_Row + 1).Value = Me.TextBox13.Value
    sh.Range("O" & last_Row + 1).Value = Now
    '----------------------------------------------------------------
    Me.TextBox1.Value = ""
    Me.TextBox2.Value = ""
    Me.TextBox3.Value = ""
    Me.TextBox4.Value = ""
    Me.TextBox5.Value = ""
    Me.TextBox6.Value = ""
    Me.TextBox7.Value = ""
    Me.TextBox8.Value = ""
    Me.TextBox9.Value = ""
    Me.TextBox10.Value = ""
    Me.TextBox11.Value = ""
    Me.TextBox12.Value = ""
    Me.TextBox13.Value = ""
    '----------------------------------
    Call Refresh_Data

End Sub


'Boton para actualizar datos
'Selected-Row

Private Sub CommandButton3_Click()

    If Me.TextBox14.Value = "" Then
        MsgBox "Seleccione el registro a actualizar"
        Exit Sub
    End If
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("Data")
    Dim Selected_Row As Long
    'Clng funcion que convierte un valor en tipo long
    Selected_Row = Application.WorksheetFunction.Match(CLng(Me.TextBox14.Value), sh.Range("A:A"), 0)
    'Validations-----------------------------------------------------
    If Me.TextBox1.Value = "" Then
    MsgBox "Por favor introduce la Fecha de Recibo", vbCritical
    Exit Sub
    End If
    If Me.TextBox2.Value = "" Then
    MsgBox "Por favor introduce el Número de Lote", vbCritical
    Exit Sub
    End If
    If Me.TextBox3.Value = "" Then
    MsgBox "Por favor introduce el Código SAP del rollo", vbCritical
    Exit Sub
    End If
    If Me.TextBox4.Value = "" Then
    MsgBox "Por favor introduce el motivo del rechazo", vbCritical
    Exit Sub
    End If
    If Me.TextBox5.Value = "" Then
    MsgBox "Por favor introduce la clave HC", vbCritical
    Exit Sub
    End If
    If Me.TextBox6.Value = "" Then
    MsgBox "Por favor introduce la Descripcion del HC", vbCritical
    Exit Sub
    End If
    If Me.TextBox7.Value = "" Then
    MsgBox "Por favor introduce el número de pedido", vbCritical
    Exit Sub
    End If
    If Me.TextBox8.Value = "" Then
    MsgBox "Por favor introduce la Fecha de Rechazo", vbCritical
    Exit Sub
    End If
    If Me.TextBox9.Value = "" Then
    MsgBox "Por favor introduce la Cantidad Rechazada", vbCritical
    Exit Sub
    End If
    If Me.TextBox10.Value = "" Then
    MsgBox "Por favor introduce el nombre del Proveedor del Metal", vbCritical
    Exit Sub
    End If
    If Me.TextBox11.Value = "" Then
    MsgBox "Por favor introduce descripción del código SAP", vbCritical
    Exit Sub
    End If
    If Me.TextBox12.Value = "" Then
    MsgBox "Por favor introduce el número de Factura/Remisión", vbCritical
    Exit Sub
    End If
    If Me.TextBox13.Value = "" Then
    MsgBox "Por favor introduce tú nombre!", vbCritical
    Exit Sub
    End If
    '-----------------------------------------------------------------
    
    sh.Range("B" & Selected_Row).Value = Me.TextBox8.Value
    sh.Range("C" & Selected_Row).Value = Me.TextBox1.Value
    sh.Range("D" & Selected_Row).Value = Me.TextBox3.Value
    sh.Range("E" & Selected_Row).Value = Me.TextBox11.Value
    sh.Range("F" & Selected_Row).Value = Me.TextBox2.Value
    sh.Range("G" & Selected_Row).Value = Me.TextBox9.Value
    sh.Range("H" & Selected_Row).Value = Me.TextBox4.Value
    sh.Range("I" & Selected_Row).Value = Me.TextBox10.Value
    sh.Range("J" & Selected_Row).Value = Me.TextBox7.Value
    sh.Range("K" & Selected_Row).Value = Me.TextBox12.Value
    sh.Range("L" & Selected_Row).Value = Me.TextBox5.Value
    sh.Range("M" & Selected_Row).Value = Me.TextBox6.Value
    sh.Range("N" & Selected_Row).Value = Me.TextBox13.Value
    sh.Range("O" & Selected_Row).Value = Now
    '----------------------------------------------------------------
    Me.TextBox1.Value = ""
    Me.TextBox2.Value = ""
    Me.TextBox3.Value = ""
    Me.TextBox4.Value = ""
    Me.TextBox5.Value = ""
    Me.TextBox6.Value = ""
    Me.TextBox7.Value = ""
    Me.TextBox8.Value = ""
    Me.TextBox9.Value = ""
    Me.TextBox10.Value = ""
    Me.TextBox11.Value = ""
    Me.TextBox12.Value = ""
    Me.TextBox13.Value = ""
    Me.TextBox14.Value = ""
    '----------------------------------
    Call Refresh_Data


End Sub

'boton para borrar registro
Private Sub CommandButton4_Click()
    If Me.TextBox14.Value = "" Then
        MsgBox "Seleccione el registro a borrar"
        Exit Sub
    End If
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("Data")
    Dim Selected_Row As Long
    'Clng funcion que convierte un valor en tipo long
    Selected_Row = Application.WorksheetFunction.Match(CLng(Me.TextBox14.Value), sh.Range("A:A"), 0)
    '----------------------------------------------
    sh.Range("A" & Selected_Row).EntireRow.Delete
    '----------------------------------------------
    Me.TextBox1.Value = ""
    Me.TextBox2.Value = ""
    Me.TextBox3.Value = ""
    Me.TextBox4.Value = ""
    Me.TextBox5.Value = ""
    Me.TextBox6.Value = ""
    Me.TextBox7.Value = ""
    Me.TextBox8.Value = ""
    Me.TextBox9.Value = ""
    Me.TextBox10.Value = ""
    Me.TextBox11.Value = ""
    Me.TextBox12.Value = ""
    Me.TextBox13.Value = ""
    Me.TextBox14.Value = ""
    
    Call Refresh_Data

End Sub

Private Sub CommandButton5_Click()
    ThisWorkbook.Save
    MsgBox "Archivo Guardado!"
End Sub
Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'rellenar el formulario al dar doble click sobre los datos capturados
    Me.TextBox14.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 0)
    Me.TextBox1.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 2)
    Me.TextBox2.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 5)
    Me.TextBox3.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 3)
    Me.TextBox4.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 7)
    Me.TextBox5.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 11)
    Me.TextBox6.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 12)
    Me.TextBox7.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 9)
    Me.TextBox8.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 1)
    Me.TextBox9.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 6)
    Me.TextBox10.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 8)
    Me.TextBox11.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 4)
    Me.TextBox12.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 10)
    Me.TextBox13.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 13)
    
End Sub


Private Sub TextBox1_Change()
    
    If Bandera = False Then
    
        If Len(TextBox1.Value) > 10 Then
            TextBox1.Value = Mid(TextBox1.Value, 1, 10)
            MsgBox "El formato de la fecha es dd/mm/aaaa"
        Else
        
            If Len(TextBox1.Value) = 2 Then
                TextBox1.Value = TextBox1.Value & "/"
            End If
            
            If Len(TextBox1.Value) = 5 Then
                TextBox1.Value = TextBox1.Value & "/"
            End If
        End If
        
    End If
End Sub

Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    If KeyCode = 8 Then
        Bandera = True
    Else
        Bandera = False
    End If

End Sub

Private Sub TextBox8_Change()
'forzar formato de fecha
    If Bandera = False Then
    
        If Len(TextBox8.Value) > 10 Then
            TextBox8.Value = Mid(TextBox8.Value, 1, 10)
            MsgBox "El formato de la fecha es dd/mm/aaaa"
        Else
        
            If Len(TextBox8.Value) = 2 Then
                TextBox8.Value = TextBox8.Value & "/"
            End If
            
            If Len(TextBox8.Value) = 5 Then
                TextBox8.Value = TextBox8.Value & "/"
            End If
        End If
        
    End If
End Sub

Private Sub UserForm_Activate()
    Call Refresh_Data
End Sub

Sub Refresh_Data()
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("Data")
    Dim last_Row As Long
    last_Row = Application.WorksheetFunction.CountA(sh.Range("A:A"))
    
    With Me.ListBox1
        .ColumnHeads = True
        'numero de columnas
        .ColumnCount = 15
        'ancho de las columnas
        .ColumnWidths = "30,65"
        '70,100,70,50,200,65,65,65,65,200,65,65"
        
        'de donde tomar el titulo de las columnas
        If last_Row = 1 Then
        .RowSource = "Data!A2:O2"
        Else
        .RowSource = "Data!A2:O" & last_Row
        End If
        
    End With
        
End Sub
