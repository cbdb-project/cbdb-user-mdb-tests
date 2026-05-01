Option Compare Database
Public gDisplayLanguage As String, gLabelsOK As Boolean

Public Sub changeDisplayLanguage()
    Dim tLabelLanguage(3, 3) As String, tLang As Integer
    
    Dim tRstLabelList As DAO.Recordset, ti As Integer
    
    Set tRstLabelList = CurrentDb.OpenRecordset("FormLabels", dbOpenTable)
    
    tRstLabelList.Index = "label"
    
    gLabelsOK = False
    With tRstLabelList
        .MoveFirst
        ti = 1
        
        Do While ti < 3 And Not .EOF
            If !c_form = "SF_POST" Then
                gLabelsOK = True
                If ti <> !c_label_id Then
                    MsgBox "Uh oh:  mismatched label table"
                    gLabelsOK = False
                    Exit Do
                End If
                tLabelLanguage(1, ti) = !c_english
                tLabelLanguage(2, ti) = !c_fanti
                tLabelLanguage(3, ti) = !c_jianti
                ti = ti + 1
            End If
            .MoveNext
        Loop
    End With
    ' tRstLabelList.Close
    Set tRstLabelList = Nothing
    
    If gLabelsOK Then
        If gDisplayLanguage = "E" Then
            tLang = 1
        ElseIf gDisplayLanguage = "T" Then
            tLang = 2
        Else
            tLang = 3
        End If
        '
        '  now comes the basic routine
        '
        ' Me.CmdAddNew.Caption = tLabelLanguage(tLang, 1)
        ' Me.CmdDelete.Caption = tLabelLanguage(tLang, 2)
        
        ' Me.POST_ADDR_Subform.Form.gDisplayLanguage = gDisplayLanguage
        ' Me.POST_ADDR_Subform.Form.changeDisplayLanguage
        
        'MsgBox "Beginning Posted_to_Office_data_2 subform 1"
        Me.POSTED_TO_OFFICE_DATA_2_Subform.Form.gDisplayLanguage = gDisplayLanguage
        'MsgBox "Beginning Posted_to_Office_data_2 subform 2"
        Me.POSTED_TO_OFFICE_DATA_2_Subform.Form.changeDisplayLanguage
    End If
    
End Sub
