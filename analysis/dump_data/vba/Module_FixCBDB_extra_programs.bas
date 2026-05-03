Public gDisplay As String
Option Compare Database

Sub add_kin_name()
    Dim rstKIN As ADODB.Recordset
    Dim rstBIOG As ADODB.Recordset
    Dim tID As Long
    
    Set rstKIN = New ADODB.Recordset
    Set rstBIOG = New ADODB.Recordset
    
    rstKIN.Open "kin_data", CurrentProject.Connection, adOpenDynamic, _
        adLockOptimistic
    
    rstBIOG.Open "BIOG_MAIN", CurrentProject.Connection, adOpenDynamic, _
        adLockOptimistic
    
    With rstKIN
        .MoveFirst
        .Find ("c_personid = 32534")
        ' .MoveNext
        Do While Not .EOF
            '
            If .EOF Then
                Exit Do
            End If
            '
            tID = .Fields("c_kin_id")
            rstBIOG.MoveFirst
            rstBIOG.Find ("c_personid = " & Str(tID))
            If Not rstBIOG.EOF Then
                .Fields("c_kin_name") = rstBIOG.Fields("c_name")
                .Fields("c_kin_name_chn") = rstBIOG.Fields("c_name_chn")
                .Update
            End If
            .MoveNext
        Loop
    End With
    rstKIN.Close
    rstBIOG.Close
    Set rstKIN = Nothing
    Set rstBIOG = Nothing
End Sub
Sub add_assoc_name()
    Dim rstAssoc As ADODB.Recordset
    Dim rstBIOG As ADODB.Recordset
    Dim tID As Long
    
    Set rstAssoc = New ADODB.Recordset
    Set rstBIOG = New ADODB.Recordset
    
    rstAssoc.Open "assoc_data", CurrentProject.Connection, adOpenDynamic, _
        adLockOptimistic
    
    rstBIOG.Open "BIOG_MAIN", CurrentProject.Connection, adOpenDynamic, _
        adLockOptimistic
    
    With rstAssoc
        .MoveFirst
        '.Find ("c_personid = 32534")
        ' .MoveNext
        Do While Not .EOF
            '
            If .EOF Then
                Exit Do
            End If
            '
            tID = .Fields("c_assoc_id")
            rstBIOG.MoveFirst
            rstBIOG.Find ("c_personid = " & Str(tID))
            If Not rstBIOG.EOF Then
                .Fields("c_assoc_name") = rstBIOG.Fields("c_name")
                .Fields("c_assoc_name_chn") = rstBIOG.Fields("c_name_chn")
                .Update
            End If
            .MoveNext
        Loop
    End With
    rstAssoc.Close
    rstBIOG.Close
    Set rstAssoc = Nothing
    Set rstBIOG = Nothing
End Sub