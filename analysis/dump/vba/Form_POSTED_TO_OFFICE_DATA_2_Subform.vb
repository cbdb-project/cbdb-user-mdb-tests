Option Compare Database
Public gDisplayLanguage As String, gLabelsOK As Boolean


Public Sub changeDisplayLanguage()
    Dim tLabelLanguage(3, 23) As String, tLang As Integer
    
    Dim tRstLabelList As DAO.Recordset, ti As Integer
    
    Set tRstLabelList = CurrentDb.OpenRecordset("FormLabels", dbOpenTable)
    
    tRstLabelList.Index = "label"
    
    ' set the language
    Dim tmli As MsoLanguageID
    ' get the labels
    tmli = Application.LanguageSettings.LanguageID(msoLanguageIDUI)
    If tmli = msoLanguageIDSimplifiedChinese Then
        gDisplayLanguage = "S"
    ElseIf tmli = msoLanguageIDTraditionalChinese Then
        gDisplayLanguage = "T"
    ElseIf tmli = msoLanguageIDEnglishUS Then
        gDisplayLanguage = "E"
    Else
        gDisplayLanguage = "E"
    End If

    
    'MsgBox "Beginning Posted to Office Data labels"
    gLabelsOK = False
    With tRstLabelList
        .MoveFirst
        ti = 1
        
        Do While ti < 23 And Not .EOF
            If !c_form = "SF_POFF" Then
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
        'MsgBox "Beginning to change Posted to Office Data labels"
        Me.LblSequence.Caption = tLabelLanguage(tLang, 1)
        Me.LblOfficeCat.Caption = tLabelLanguage(tLang, 2)
        Me.LblPostCat.Caption = tLabelLanguage(tLang, 3)
        Me.LblFirstyear.Caption = tLabelLanguage(tLang, 4)
        Me.LblFYRange.Caption = tLabelLanguage(tLang, 5)
        Me.LblAssume.Caption = tLabelLanguage(tLang, 6)
        Me.LblFYNHYear.Caption = tLabelLanguage(tLang, 7)
        Me.LblFYIintercalary.Caption = tLabelLanguage(tLang, 8)
        Me.LblFYGZ.Caption = tLabelLanguage(tLang, 9)
        Me.LblLastyear.Caption = tLabelLanguage(tLang, 10)
        Me.LblLYRange.Caption = tLabelLanguage(tLang, 11)
        Me.LblLYNHyear.Caption = tLabelLanguage(tLang, 12)
        Me.LblLYIntercalary.Caption = tLabelLanguage(tLang, 13)
        Me.LblLYGZ.Caption = tLabelLanguage(tLang, 14)
        Me.LblPages.Caption = tLabelLanguage(tLang, 15)
        Me.LblNotes.Caption = tLabelLanguage(tLang, 16)
        Me.LblOfficeName.Caption = tLabelLanguage(tLang, 17)
        Me.LblFYNH.Caption = tLabelLanguage(tLang, 18)
        Me.LblLYNH.Caption = tLabelLanguage(tLang, 19)
        Me.LblSource.Caption = tLabelLanguage(tLang, 20)
        ' Me.CmdAddNew.Caption = tLabelLanguage(tLang, 21)
        ' Me.CmdDelete.Caption = tLabelLanguage(tLang, 22)
        
        ' Me.POSTED_TO_ADDR_DATA_2_Subform.Form.gTest = 1
        '
        ' Me.POSTED_TO_ADDR_DATA_2_Subform.Form.gDisplayLanguage = gDisplayLanguage
        ' Me.POSTED_TO_ADDR_DATA_2_Subform.Form.changeDisplayLanguage
    End If
    
End Sub

Public Sub noEdits()

    Me.AllowAdditions = False
    Me.AllowDeletions = False
    Me.AllowEdits = False
    
End Sub
