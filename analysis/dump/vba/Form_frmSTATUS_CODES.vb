Option Compare Database

Private Sub c_status_code_Click()
    Dim tRst As DAO.Recordset
    
    Set tRst = Me.Recordset
    Forms!frmPickStatus2.Form!TxtStatusID.Value = tRst!c_status_code
    Forms!frmPickStatus2.Form!TxtStatusDesc.Value = tRst!c_status_desc
    Forms!frmPickStatus2.Form!TxtStatusDescChn.Value = tRst!c_status_desc_chn
    Forms!frmPickStatus2.Form!TxtTypeID.Value = "000"
    Forms!frmPickStatus2.Form!TxtTypeDescChn.Value = ""
    Forms!frmPickStatus2.Form!TxtTypeDesc.Value = ""
    Me.DatasheetBackColor = RGB(255, 255, 255)
    Me.DatasheetForeColor = RGB(0, 0, 0)
    Forms!frmPickStatus2.Form!CmdSelectAll.Caption = "Select All"
    Forms!frmPickStatus2.Form!CmdSelectAll.Enabled = True
    Forms!frmPickStatus2.Form!CmdSelect.Enabled = True
    Set tRst = Nothing

End Sub

Private Sub c_status_desc_chn_Click()
    Dim tRst As DAO.Recordset
    
    Set tRst = Me.Recordset
    Forms!frmPickStatus2.Form!TxtStatusID.Value = tRst!c_status_code
    Forms!frmPickStatus2.Form!TxtStatusDesc.Value = tRst!c_status_desc
    Forms!frmPickStatus2.Form!TxtStatusDescChn.Value = tRst!c_status_desc_chn
    Forms!frmPickStatus2.Form!TxtTypeID.Value = "000"
    Forms!frmPickStatus2.Form!TxtTypeDescChn.Value = ""
    Forms!frmPickStatus2.Form!TxtTypeDesc.Value = ""
    Me.DatasheetBackColor = RGB(255, 255, 255)
    Me.DatasheetForeColor = RGB(0, 0, 0)
    Forms!frmPickStatus2.Form!CmdSelectAll.Caption = "Select All"
    Forms!frmPickStatus2.Form!CmdSelectAll.Enabled = True
    Forms!frmPickStatus2.Form!CmdSelect.Enabled = True
    Set tRst = Nothing

End Sub

Private Sub c_status_desc_Click()
    Dim tRst As DAO.Recordset
    
    Set tRst = Me.Recordset
    Forms!frmPickStatus2.Form!TxtStatusID.Value = tRst!c_status_code
    Forms!frmPickStatus2.Form!TxtStatusDesc.Value = tRst!c_status_desc
    Forms!frmPickStatus2.Form!TxtStatusDescChn.Value = tRst!c_status_desc_chn
    Forms!frmPickStatus2.Form!TxtTypeID.Value = "000"
    Forms!frmPickStatus2.Form!TxtTypeDesc.Value = ""
    Forms!frmPickStatus2.Form!TxtTypeDescChn.Value = ""
    Me.DatasheetBackColor = RGB(255, 255, 255)
    Me.DatasheetForeColor = RGB(0, 0, 0)
    Forms!frmPickStatus2.Form!CmdSelectAll.Caption = "Select All"
    Forms!frmPickStatus2.Form!CmdSelectAll.Enabled = True
    Forms!frmPickStatus2.Form!CmdSelect.Enabled = True
    Set tRst = Nothing

End Sub
