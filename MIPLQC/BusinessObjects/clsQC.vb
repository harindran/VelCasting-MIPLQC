Public Class clsQC
    Public Const Formtype = "MIPLQC"
    Dim objForm As SAPbouiCOM.Form
    Dim strSQL As String, strQuery As String
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim QCHeader As SAPbouiCOM.DBDataSource
    Dim QCLine As SAPbouiCOM.DBDataSource
    Dim InWhse As String
    Dim AccWhse As String
    Dim RejWhse As String
    Dim RewWhse As String
    Dim objRecordSet As SAPbobsCOM.Recordset
    Dim AccQty, RejQty, RewQty As Integer
    Dim TotQty, PendQty, InspQty As Integer
    Dim objCombo As SAPbouiCOM.ComboBox
    Public Sub LoadScreen()
        objForm = objAddOn.objUIXml.LoadScreenXML("QC.xml", Mukesh.SBOLib.UIXML.enuResourceType.Embeded, Formtype)
        objForm.PaneLevel = 1
        BranchEnabled(objForm.UniqueID)
        LoadSeries(objForm.UniqueID)
        objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
        objForm.Items.Item("13").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
        objForm.Items.Item("13B").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
        objForm.Items.Item("15").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
        objForm.Items.Item("15B").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
        objForm.Items.Item("23").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
        objForm.Items.Item("23B").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
        objForm.Items.Item("51").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
        objForm.Items.Item("20").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
        objForm.Items.Item("34").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
        'objForm.Items.Item("6").Specific.Active = True

    End Sub
    Public Sub ItemEvent(FormUID As String, pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        Dim objcombo As SAPbouiCOM.ComboBox
        objcombo = objForm.Items.Item("8").Specific
        If pVal.BeforeAction = True Then
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    'If pVal.ItemUID = "1" And (objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then
                    '    If Not validate(FormUID) Then
                    '        BubbleEvent = False
                    '        Exit Sub
                    '        'Else
                    '        '    objAddOn.objCompany.StartTransaction()
                    '        '    If StockPosting_GRN(FormUID) Then
                    '        '        objAddOn.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    '        '    Else
                    '        '        BubbleEvent = False
                    '        '        objAddOn.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    '        '        Exit Sub
                    '        '    End If
                    '    End If
                    'End If
                    If pVal.ItemUID = "1" And objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        If objForm.Items.Item("34").Specific.String = "" Or objForm.Items.Item("34").Specific.String = "0" Then
                            objForm.Items.Item("34").Specific.String = getQCEntry(FormUID)
                        Else
                            BubbleEvent = False
                            Exit Sub
                        End If

                    End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
                    If objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        If objForm.Items.Item("34").Specific.string <> "" Then
                            objAddOn.objApplication.MessageBox("Please update")
                            BubbleEvent = False
                        End If
                    End If
            End Select
        Else
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    If pVal.ItemUID = "2A" And objForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        StockTransfer_BinLocation(FormUID)
                        ' BatchUpdate()
                    ElseIf pVal.ItemUID = "1000001" Then
                        objForm.PaneLevel = 1
                    ElseIf pVal.ItemUID = "32" Then
                        objForm.PaneLevel = 2
                    End If

                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    If (pVal.ItemUID = "13") Or (pVal.ItemUID = "23") Or (pVal.ItemUID = "15") Or (pVal.ItemUID = "51" And objcombo.Selected.Value = "R") Then
                        getDocumentEntry(FormUID, objcombo.Selected.Value)
                        If pVal.ItemUID <> "15" Then
                            LoadInspectionMatrix(FormUID, objForm.Items.Item("8").Specific.selected.value)
                        End If
                    ElseIf (pVal.ItemUID = "51" And (objcombo.Selected.Value = "P")) Then
                        LoadInspectionMatrix(FormUID, objForm.Items.Item("8").Specific.selected.value)
                    ElseIf pVal.ItemUID = "20" And pVal.ColUID = "7D" Then
                        objMatrix.Columns.Item("7E").Cells.Item(pVal.Row).Specific.string = CDbl(objMatrix.Columns.Item("7").Cells.Item(pVal.Row).Specific.string) * CDbl(objMatrix.Columns.Item("7D").Cells.Item(pVal.Row).Specific.string)
                    ElseIf pVal.ItemUID = "20" And pVal.ColUID = "7AA" Then
                        objAddOn.objApplication.Menus.Item("1300").Activate()
                    End If
                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    If pVal.ItemUID = "8" Then
                        TypeSelection(FormUID)
                    ElseIf pVal.ItemUID = "21" Then
                        objForm.Items.Item("4").Specific.String = objForm.BusinessObject.GetNextSerialNumber(objForm.Items.Item("21").Specific.selected.value, Formtype)
                    ElseIf pVal.ItemUID = "50" Then
                        CFLConditions(FormUID)
                    ElseIf pVal.ItemUID = "20" And pVal.ColUID = "7AA" Then
                        objMatrix = objForm.Items.Item("20").Specific
                        objcombo = objMatrix.Columns.Item("7AA").Cells.Item(pVal.Row).Specific
                        If objcombo.Selected.Value <> "-" And CInt(objMatrix.Columns.Item("7").Cells.Item(pVal.Row).Specific.String) = 0 Then
                            objAddOn.objApplication.SetStatusBarMessage("Please update the Rejected Qty & then Select Rejected Location... Line: " & pVal.Row, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                            objcombo.Select("-", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            objMatrix.Columns.Item("7").Cells.Item(pVal.Row).Click()
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                    'If pVal.ItemUID = "20" And pVal.ColUID = "7_1" And pVal.CharPressed = Asc(LCase("d")) Then
                    '    objAddOn.objRejDetails.LoadScreen(FormUID, pVal.Row)
                    'End If
                    objMatrix = objForm.Items.Item("20").Specific
                    If pVal.ItemUID = "20" And pVal.ColUID = "7_1" And (pVal.CharPressed = 37 Or pVal.CharPressed = 39 Or pVal.CharPressed = 9) Then
                        If objMatrix.Columns.Item("7_1").Cells.Item(pVal.Row).Specific.string = "" And CDbl(objMatrix.Columns.Item("7").Cells.Item(pVal.Row).Specific.string) > 0 Then
                            objAddOn.objRejDetails.LoadScreen(FormUID, pVal.Row)
                        Else
                            Exit Sub
                        End If
                    End If
                    Dim ColID As Integer = objMatrix.GetCellFocus().ColumnIndex
                    If pVal.ItemUID = "20" And pVal.CharPressed = 38 And pVal.Modifiers = SAPbouiCOM.BoModifiersEnum.mt_None Then  'up
                        objMatrix.SetCellFocus(pVal.Row - 1, ColID)
                        objMatrix.SelectRow(pVal.Row - 1, True, False)
                    ElseIf pVal.ItemUID = "20" And pVal.CharPressed = 40 And pVal.Modifiers = SAPbouiCOM.BoModifiersEnum.mt_None Then 'down
                        objMatrix.SetCellFocus(pVal.Row + 1, ColID)
                        objMatrix.SelectRow(pVal.Row + 1, True, False)
                    ElseIf pVal.ItemUID = "20" And pVal.CharPressed = 37 And pVal.Modifiers = SAPbouiCOM.BoModifiersEnum.mt_None Then 'Left
                        objMatrix.SetCellFocus(pVal.Row, ColID - 1)
                    ElseIf pVal.ItemUID = "20" And pVal.CharPressed = 39 And pVal.Modifiers = SAPbouiCOM.BoModifiersEnum.mt_None Then 'Right
                        objMatrix.SetCellFocus(pVal.Row, ColID + 1)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                    If pVal.ItemUID = "20" And pVal.ColUID = "7_1" Then
                        objMatrix = objForm.Items.Item("20").Specific
                        objAddOn.objRejDetails.LoadScreen(FormUID, pVal.Row, objMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.String)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    If pVal.ItemUID = "20" And (pVal.ColUID = "6A" Or pVal.ColUID = "7A" Or pVal.ColUID = "8A") Then
                        CFL(FormUID, pVal)
                    End If
            End Select
        End If
    End Sub

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objForm = objAddOn.objApplication.Forms.Item(BusinessObjectInfo.FormUID)
            If BusinessObjectInfo.BeforeAction = True Then
                Select Case BusinessObjectInfo.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD, SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                            'If BusinessObjectInfo.BeforeAction = True Then
                            If validate(objForm.UniqueID) = False Then
                                System.Media.SystemSounds.Asterisk.Play()
                                BubbleEvent = False
                                Exit Sub
                            Else

                            End If
                            'If objForm.Items.Item("34").Specific.string = "" Then
                            '    System.Media.SystemSounds.Asterisk.Play()
                            '    objAddOn.objApplication.SetStatusBarMessage("Stock Not Transferred!!!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            '    BubbleEvent = False
                            '    Exit Sub
                            'End If
                            'End If
                        End If

                        'Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                        '    'If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        '    objForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                        '    objForm.Items.Item("BtnInv").Enabled = True
                        '    objForm.EnableMenu("1282", True)
                        '    'End If
                End Select
            Else

            End If

        Catch ex As Exception
            objAddOn.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub

    Private Sub CFLConditions(ByVal FormUID As String)
        setCFLCond(FormUID, "CFL_1")
        setCFLCond(FormUID, "CFL_2")
        setCFLCond(FormUID, "CFL_3")
    End Sub
    Private Sub setCFLCond(ByVal FormUID As String, ByVal CFLId As String)
        Dim objCFL As SAPbouiCOM.ChooseFromList
        Dim objCondition As SAPbouiCOM.Condition
        Dim objConditions As SAPbouiCOM.Conditions
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objCombo = objForm.Items.Item("50").Specific
        objCFL = objForm.ChooseFromLists.Item(CFLId)
        For i As Integer = 0 To objCFL.GetConditions.Count - 1
            objCFL.SetConditions(Nothing)
        Next

        objConditions = objCFL.GetConditions()
        objCondition = objConditions.Add()
        objCondition.Alias = "BPLid"
        objCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
        objCondition.CondVal = objCombo.Selected.Value
        objCFL.SetConditions(objConditions)
    End Sub
    Public Sub LoadSeries(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        QCHeader = objForm.DataSources.DBDataSources.Item("@MIPLQC")
        QCLine = objForm.DataSources.DBDataSources.Item("@MIPLQC1")
        objForm.Items.Item("21").Specific.validvalues.loadseries(Formtype, SAPbouiCOM.BoSeriesMode.sf_Add)
        objForm.Items.Item("21").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        QCHeader.SetValue("DocNum", 0, objAddOn.objGenFunc.GetDocNum(Formtype, CInt(objForm.Items.Item("21").Specific.Selected.value)))
        QCHeader.SetValue("DocEntry", 0, objAddOn.objGenFunc.GetNextDocEntry_Value("@MIPLQC", CInt(objForm.Items.Item("21").Specific.Selected.value)))
        objForm.Items.Item("6").Specific.String = "A"
        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            objCombo = objForm.Items.Item("8").Specific
            objCombo.Select("G", SAPbouiCOM.BoSearchKey.psk_ByValue)
        End If
        objCombo = objForm.Items.Item("50").Specific
        If objCombo.ValidValues.Count = 0 Then
            If objAddOn.HANA Then
                strSQL = "select ""BPLId"", ""BPLName"" from OBPL"
            Else
                strSQL = "select BPLId, BPLName from OBPL"
            End If

            objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objRecordSet.DoQuery(strSQL)
            While Not objRecordSet.EoF
                objCombo.ValidValues.Add(objRecordSet.Fields.Item(0).Value, objRecordSet.Fields.Item(1).Value)
                objRecordSet.MoveNext()
            End While
            objRecordSet = Nothing
        End If
    End Sub
    Private Sub BranchEnabled(ByVal FormUID As String)

        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        If objAddOn.HANA Then
            strSQL = "SELECT ""MltpBrnchs"" FROM OADM"

        Else
            strSQL = "SELECT MltpBrnchs FROM OADM"
        End If
        objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objRecordSet.DoQuery(strSQL)
        If objRecordSet.EoF Then Exit Sub
        If UCase(Trim(CStr(objRecordSet.Fields.Item("MltpBrnchs").Value))) = "Y" Then

            objForm.Items.Item("50").Visible = True
        Else
            objForm.Items.Item("50").Visible = False
        End If

    End Sub
    
    Private Sub CFL(ByVal FormUID As String, ByVal pval As SAPbouiCOM.ItemEvent)
        Try
            Dim objCFLEvent As SAPbouiCOM.ChooseFromListEvent
            Dim objDataTable As SAPbouiCOM.DataTable
            objCFLEvent = pval
            objDataTable = objCFLEvent.SelectedObjects
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            objMatrix = objForm.Items.Item("20").Specific
            Select Case objCFLEvent.ChooseFromListUID
                Case "CFL_1"
                    Try
                        If objDataTable Is Nothing Then
                        Else
                            objMatrix.Columns.Item("6A").Cells.Item(pval.Row).Specific.String = objDataTable.GetValue("WhsCode", 0)

                        End If
                    Catch ex As Exception
                        objMatrix.Columns.Item("6A").Cells.Item(pval.Row).Specific.String = objDataTable.GetValue("WhsCode", 0)
                    End Try
                Case "CFL_2"
                    Try
                        If objDataTable Is Nothing Then
                        Else
                            objMatrix.Columns.Item("7A").Cells.Item(pval.Row).Specific.String = objDataTable.GetValue("WhsCode", 0)

                        End If
                    Catch ex As Exception
                        objMatrix.Columns.Item("7A").Cells.Item(pval.Row).Specific.String = objDataTable.GetValue("WhsCode", 0)
                    End Try
                Case "CFL_3"
                    Try
                        If objDataTable Is Nothing Then
                        Else
                            objMatrix.Columns.Item("8A").Cells.Item(pval.Row).Specific.String = objDataTable.GetValue("WhsCode", 0)
                        End If
                    Catch ex As Exception
                        objMatrix.Columns.Item("8A").Cells.Item(pval.Row).Specific.String = objDataTable.GetValue("WhsCode", 0)
                    End Try
            End Select
          
        Catch ex As Exception
            'MsgBox("Error : " + ex.Message + vbCrLf + "Position : " + ex.StackTrace, MsgBoxStyle.Critical)
        End Try

    End Sub
    Private Function InspectedQtyCorrect(ByVal FormUID As String, ByVal Row As Integer) As Boolean
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("20").Specific
        If objMatrix.RowCount = 0 Then
            objAddOn.objApplication.SetStatusBarMessage("Inspection Matrix is Empty", SAPbouiCOM.BoMessageTime.bmt_Short)
            Return False
        End If
        Try
            TotQty = IIf(objMatrix.Columns.Item("3").Cells.Item(Row).Specific.string.trim = "", 0, CInt(objMatrix.Columns.Item("3").Cells.Item(Row).Specific.string))

            InspQty = IIf(objMatrix.Columns.Item("4").Cells.Item(Row).Specific.string.Trim = "", 0, CInt(objMatrix.Columns.Item("4").Cells.Item(Row).Specific.string))

            PendQty = IIf(objMatrix.Columns.Item("5").Cells.Item(Row).Specific.string.Trim = "", 0, CInt(objMatrix.Columns.Item("5").Cells.Item(Row).Specific.string))

            AccQty = IIf(objMatrix.Columns.Item("6").Cells.Item(Row).Specific.string.Trim = "", 0, CInt(objMatrix.Columns.Item("6").Cells.Item(Row).Specific.string))

            RejQty = IIf(objMatrix.Columns.Item("7").Cells.Item(Row).Specific.string.Trim = "", 0, CInt(objMatrix.Columns.Item("7").Cells.Item(Row).Specific.string))

            RewQty = IIf(objMatrix.Columns.Item("8").Cells.Item(Row).Specific.string.Trim = "", 0, CInt(objMatrix.Columns.Item("8").Cells.Item(Row).Specific.string))
            Dim TTQty As Integer = AccQty + RejQty + RewQty
            If PendQty < TTQty Then
                objAddOn.objApplication.SetStatusBarMessage("Please check quantity")
                Return False
            End If
            If (CInt(objMatrix.Columns.Item("9").Cells.Item(Row).Specific.String) <> TTQty) Then
                objAddOn.objApplication.SetStatusBarMessage("Please check the QtyInspected in Line No :" & CStr(Row))
                Return False
            End If
            If CInt(objMatrix.Columns.Item("7").Cells.Item(Row).Specific.string) > 0 Then
                objMatrix.Columns.Item("7B").Cells.Item(Row).Specific.string = CStr((CInt(objMatrix.Columns.Item("7").Cells.Item(Row).Specific.string) / TTQty) * 100)
                objMatrix.Columns.Item("7C").Cells.Item(Row).Specific.string = CDbl(objMatrix.Columns.Item("7B").Cells.Item(Row).Specific.string) * 10000
                objMatrix.Columns.Item("7E").Cells.Item(Row).Specific.string = CDbl(objMatrix.Columns.Item("7").Cells.Item(Row).Specific.string) * CDbl(objMatrix.Columns.Item("7D").Cells.Item(Row).Specific.string)
            End If
            Return True
        Catch ex As Exception
            objAddOn.objApplication.SetStatusBarMessage(ex.Message)
            Return False
        End Try
    End Function
    Public Sub TypeSelection(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        Dim objCombo As SAPbouiCOM.ComboBox
        objCombo = objForm.Items.Item("8").Specific
        Select Case objCombo.Selected.Value
            Case "-"
                objForm.Items.Item("12").Visible = False
                objForm.Items.Item("13").Visible = False
                objForm.Items.Item("13A").Visible = False
                objForm.Items.Item("13B").Visible = False
                objForm.Items.Item("14").Visible = False
                objForm.Items.Item("15").Visible = False
                objForm.Items.Item("15A").Visible = False
                objForm.Items.Item("15B").Visible = False
                objForm.Items.Item("22").Visible = False
                objForm.Items.Item("23").Visible = False
                objForm.Items.Item("23A").Visible = False
                objForm.Items.Item("23B").Visible = False
                objForm.Items.Item("49").Visible = False
                objForm.Items.Item("51").Visible = False
                objForm.Items.Item("51A").Visible = False
                objForm.Items.Item("51B").Visible = False
            Case "G"
                objForm.Items.Item("12").Visible = True
                objForm.Items.Item("13").Visible = True
                objForm.Items.Item("13A").Visible = True
                objForm.Items.Item("13B").Visible = True
                objForm.Items.Item("14").Visible = False
                objForm.Items.Item("15").Visible = False
                objForm.Items.Item("15A").Visible = False
                objForm.Items.Item("15B").Visible = False
                objForm.Items.Item("22").Visible = False
                objForm.Items.Item("23").Visible = False
                objForm.Items.Item("23A").Visible = False
                objForm.Items.Item("23B").Visible = False
                objForm.Items.Item("49").Visible = False
                objForm.Items.Item("51").Visible = False
                objForm.Items.Item("51A").Visible = False
                objForm.Items.Item("51B").Visible = False
                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    objForm.Items.Item("15").Specific.String = ""
                    objForm.Items.Item("15B").Specific.String = ""
                    objForm.Items.Item("23").Specific.String = ""
                    objForm.Items.Item("23B").Specific.String = ""
                    objForm.Items.Item("51").Specific.String = ""
                    objForm.Items.Item("51B").Specific.String = ""
                End If
            Case "P"
                objForm.Items.Item("12").Visible = False
                objForm.Items.Item("13").Visible = False
                objForm.Items.Item("13A").Visible = False
                objForm.Items.Item("13B").Visible = False
                objForm.Items.Item("14").Visible = True
                objForm.Items.Item("15").Visible = True
                objForm.Items.Item("15A").Visible = True
                objForm.Items.Item("15B").Visible = True
                objForm.Items.Item("22").Visible = False
                objForm.Items.Item("23").Visible = False
                objForm.Items.Item("23A").Visible = False
                objForm.Items.Item("23B").Visible = False
                objForm.Items.Item("49").Visible = True
                objForm.Items.Item("51").Visible = True
                objForm.Items.Item("51A").Visible = True
                objForm.Items.Item("51B").Visible = True
                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    objForm.Items.Item("13").Specific.String = ""
                    objForm.Items.Item("13B").Specific.String = ""
                    objForm.Items.Item("23").Specific.String = ""
                    objForm.Items.Item("23B").Specific.String = ""
                End If
            Case "T"
                objForm.Items.Item("12").Visible = False
                objForm.Items.Item("13").Visible = False
                objForm.Items.Item("13A").Visible = False
                objForm.Items.Item("13B").Visible = False
                objForm.Items.Item("14").Visible = False
                objForm.Items.Item("15").Visible = False
                objForm.Items.Item("15A").Visible = False
                objForm.Items.Item("15B").Visible = False
                objForm.Items.Item("22").Visible = True
                objForm.Items.Item("23").Visible = True
                objForm.Items.Item("23A").Visible = True
                objForm.Items.Item("23B").Visible = True
                objForm.Items.Item("49").Visible = False
                objForm.Items.Item("51").Visible = False
                objForm.Items.Item("51A").Visible = False
                objForm.Items.Item("51B").Visible = False
                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    objForm.Items.Item("13").Specific.String = ""
                    objForm.Items.Item("13B").Specific.String = ""
                    objForm.Items.Item("15").Specific.String = ""
                    objForm.Items.Item("15B").Specific.String = ""
                    objForm.Items.Item("51").Specific.String = ""
                    objForm.Items.Item("51B").Specific.String = ""
                End If
            Case "R"
                objForm.Items.Item("12").Visible = False
                objForm.Items.Item("13").Visible = False
                objForm.Items.Item("13A").Visible = False
                objForm.Items.Item("13B").Visible = False
                objForm.Items.Item("14").Visible = False
                objForm.Items.Item("15").Visible = False
                objForm.Items.Item("15A").Visible = False
                objForm.Items.Item("15B").Visible = False
                objForm.Items.Item("22").Visible = False
                objForm.Items.Item("23").Visible = False
                objForm.Items.Item("23A").Visible = False
                objForm.Items.Item("23B").Visible = False
                objForm.Items.Item("49").Visible = True
                objForm.Items.Item("51").Visible = True
                objForm.Items.Item("51A").Visible = True
                objForm.Items.Item("51B").Visible = True
                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    objForm.Items.Item("13").Specific.String = ""
                    objForm.Items.Item("13B").Specific.String = ""
                    objForm.Items.Item("23").Specific.String = ""
                    objForm.Items.Item("23B").Specific.String = ""
                    objForm.Items.Item("15").Specific.String = ""
                    objForm.Items.Item("15B").Specific.String = ""
                End If
        End Select
        setWhse(objCombo.Selected.Value)
    End Sub
    Private Sub setWhse(ByVal doctype As String)
        'If objAddOn.HANA Then
        '    strSQL = "SELECT * FROM ""@QCWHSE"" WHERE ""U_Type"" = '" & doctype & "';"

        'Else
        '    strSQL = "select Code,Name,U_Type,U_InWhse,U_AccWhse,U_RejWhse,U_RewWhse from [@QCWHSE] where U_Type='" & doctype & "'"
        'End If
        'objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        'objRecordSet = objAddOn.objGenFunc.DoQuery(strSQL)
        'If Not objRecordSet.EoF Then
        '    InWhse = objRecordSet.Fields.Item("U_InWhse").Value
        '    AccWhse = objRecordSet.Fields.Item("U_AccWhse").Value
        '    RejWhse = objRecordSet.Fields.Item("U_RejWhse").Value
        '    RewWhse = objRecordSet.Fields.Item("U_RewWhse").Value

        'End If

    End Sub
    Public Sub LoadInspectionMatrix(ByVal FormUID As String, ByVal Type As String)
        Try
            Dim DocEntry As String = ""
            Dim objcombo As SAPbouiCOM.ComboBox
            objcombo = objForm.Items.Item("8").Specific
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            objMatrix = objForm.Items.Item("20").Specific
            objMatrix.Clear()
            Select Case objcombo.Selected.Value
                Case "P"
                    DocEntry = objForm.Items.Item("51").Specific.String
                Case "G"
                    DocEntry = objForm.Items.Item("13B").Specific.String
                Case "T"
                    DocEntry = objForm.Items.Item("23B").Specific.String
                Case "R"
                    DocEntry = objForm.Items.Item("51").Specific.String

            End Select
           


            If DocEntry = "" Then
                objAddOn.objApplication.MessageBox("Select valid DocEntry")
                Exit Sub
            End If
            strSQL = getDetailsQuery(DocEntry, Type)
            objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objRecordSet.DoQuery(strSQL)
            QCLine = objForm.DataSources.DBDataSources.Item("@MIPLQC1")
            If objRecordSet.RecordCount > 0 Then
                While Not objRecordSet.EoF
                    If CDbl(objRecordSet.Fields.Item("PendQty").Value) > 0 Then
                        If objMatrix.RowCount = 0 Then
                            objMatrix.AddRow()
                        ElseIf objMatrix.Columns.Item("1").Cells.Item(objMatrix.RowCount).Specific.String <> "" Then
                            objMatrix.AddRow()
                        End If
                        QCLine.Clear()
                        objMatrix.GetLineData(objMatrix.RowCount)
                        QCLine.SetValue("LineId", 0, objRecordSet.Fields.Item("LineId").Value)
                        QCLine.SetValue("U_BaseLinNum", 0, objRecordSet.Fields.Item("LineNum").Value)
                        QCLine.SetValue("U_ItemCode", 0, objRecordSet.Fields.Item("ItemCode").Value)
                        Dim strItemQry As String
                        Dim objItemRS As SAPbobsCOM.Recordset
                        If objAddOn.HANA Then
                            strItemQry = "SELECT T1.""ItmsGrpNam"" FROM OITM T0  INNER JOIN OITB T1 ON T0.""ItmsGrpCod"" = T1.""ItmsGrpCod"" WHERE T0.""ItemCode"" = '" & objRecordSet.Fields.Item("ItemCode").Value & "'"
                        Else
                            strItemQry = "select T1.ItmsGrpNam from OITM T0 join  OITB T1 on T0.ItmsGrpCod=T1.ItmsGrpCod where T0.ItemCode='" & objRecordSet.Fields.Item("ItemCode").Value & "'"
                        End If

                        objItemRS = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        objItemRS.DoQuery(strItemQry)
                        Dim groupname As String = objItemRS.Fields.Item("ItmsGrpNam").Value
                        objItemRS = Nothing
                        QCLine.SetValue("U_ItemName", 0, objRecordSet.Fields.Item("Dscription").Value)
                        QCLine.SetValue("U_ItGrpNm", 0, groupname)
                        '   QCLine.SetValue("U_TotQty", 0, objRecordSet.Fields.Item("TotQty").Value)
                        Dim ConvTotQty As Double = objRecordSet.Fields.Item("TotQty").Value
                        QCLine.SetValue("U_TotQty", 0, ConvTotQty)
                        QCLine.SetValue("U_InvUom", 0, objRecordSet.Fields.Item("InvntryUom").Value)
                        'QCLine.SetValue("U_InspQty", 0, objRecordSet.Fields.Item("InspQty").Value)
                        'QCLine.SetValue("U_PendQty", 0, objRecordSet.Fields.Item("PendQty").Value)
                        QCLine.SetValue("U_InspQty", 0, objRecordSet.Fields.Item("InspQty").Value)

                        QCLine.SetValue("U_PendQty", 0, objRecordSet.Fields.Item("PendQty").Value)
                        QCLine.SetValue("U_AccQty", 0, 0)
                        QCLine.SetValue("U_RejQty", 0, 0)
                        QCLine.SetValue("U_RewQty", 0, 0)
                        QCLine.SetValue("U_QtyInsp", 0, 0)
                        QCLine.SetValue("U_SmplQty", 0, 0)
                        QCLine.SetValue("U_AccWhse", 0, AccWhse)
                        QCLine.SetValue("U_RejWhse", 0, RejWhse)
                        QCLine.SetValue("U_ItemCost", 0, objRecordSet.Fields.Item("Price").Value)
                        QCLine.SetValue("U_RewWhse", 0, RewWhse)
                        objMatrix.SetLineData(objMatrix.RowCount)
                    End If
                    objRecordSet.MoveNext()
                End While
            End If

            If objRecordSet.RecordCount > 0 Then
                getItemRemovalNotify(Type)
            End If
            objAddOn.objApplication.Menus.Item("1300").Activate()
            objForm.Refresh()
            objRecordSet = Nothing
        Catch ex As Exception
            ' MsgBox(ex.ToString)
        End Try
    End Sub
    Public Function GetConvertedQty(ByVal FormUID As String, ByVal DocEntry As String, ByVal Type As String, ByVal ItemCode As String, ByVal Quantity As Double)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        Dim objMatrix As SAPbouiCOM.Matrix
        Dim objRecordSet As SAPbobsCOM.Recordset
        objMatrix = objForm.Items.Item("20").Specific
        Dim StrQuery As String = ""
        Dim convertedQty As Double = 0
        If objAddOn.HANA Then
             Select Type
                Case "G"
                    StrQuery = "select  T0. ""ItemCode"", T0.""UomCode"", T0.""UomEntry"", T1.""UgpEntry"" , T1. ""InvntryUom"" ,T2.""AltQty"" ,T2.""BaseQty"" from PDN1 T0  left outer join OITM T1 on T1.""ItemCode"" = T0.""ItemCode""" & _
                                                   "left outer join UGP1 T2 on T2.""UomEntry"" = T0.""UomEntry""  and T2.""UgpEntry"" = T1.""UgpEntry""  where  T0.""DocEntry"" ='" & DocEntry & "' and T0.""ItemCode"" ='" & ItemCode & "'"

                Case "P"
                    StrQuery = "select  T0. ""ItemCode"", T0.""UomCode"", T0.""UomEntry"", T1.""UgpEntry"" , T1. ""InvntryUom"" ,T2.""AltQty"" ,T2.""BaseQty"" from IGN1 T0  left outer join OITM T1 on T1.""ItemCode"" = T0.""ItemCode""" & _
                                                   "left outer join UGP1 T2 on T2.""UomEntry"" = T0.""UomEntry""  and T2.""UgpEntry"" = T1.""UgpEntry""  where  T0.""DocEntry"" ='" & DocEntry & "' and T0.""ItemCode"" ='" & ItemCode & "'"

                Case "T"
                    StrQuery = "select  T0. ""ItemCode"", T0.""UomCode"", T0.""UomEntry"", T1.""UgpEntry"" , T1. ""InvntryUom"" ,T2.""AltQty"" ,T2.""BaseQty"" from WTR1 T0  left outer join OITM T1 on T1.""ItemCode"" = T0.""ItemCode""" & _
                                                   "left outer join UGP1 T2 on T2.""UomEntry"" = T0.""UomEntry""  and T2.""UgpEntry"" = T1.""UgpEntry""  where  T0.""DocEntry"" ='" & DocEntry & "' and T0.""ItemCode"" ='" & ItemCode & "'"
                Case "R"
                    StrQuery = "select  T0. ""ItemCode"", T0.""UomCode"", T0.""UomEntry"", T1.""UgpEntry"" , T1. ""InvntryUom"" ,T2.""AltQty"" ,T2.""BaseQty"" from IGN1 T0  left outer join OITM T1 on T1.""ItemCode"" = T0.""ItemCode""" & _
                                                   "left outer join UGP1 T2 on T2.""UomEntry"" = T0.""UomEntry""  and T2.""UgpEntry"" = T1.""UgpEntry""  where  T0.""DocEntry"" ='" & DocEntry & "' and T0.""ItemCode"" ='" & ItemCode & "'"

            End Select
        Else
            Select Case Type
                Case "G"
                    StrQuery = "select  T0. ItemCode, T0.UomCode, T0.UoMEntry, T1.UgpEntry , T1. InvntryUom ,T2.AltQty ,T2.BaseQty from PDN1 T0  left outer join OITM T1 on T1.ItemCode = T0.ItemCode " & _
                               "left outer join UGP1 T2 on T2.UomEntry = T0.UomEntry  and T2.UgpEntry = T1.UgpEntry  where  T0.DocEntry ='" & DocEntry & "' and t0.ItemCode ='" & ItemCode & "'"

                Case "P"
                    StrQuery = "select  T0. ItemCode, T0.UomCode, T0.UoMEntry, T1.UgpEntry , T1. InvntryUom ,T2.AltQty ,T2.BaseQty from IGN1 T0  left outer join OITM T1 on T1.ItemCode = T0.ItemCode " & _
                               "left outer join UGP1 T2 on T2.UomEntry = T0.UomEntry  and T2.UgpEntry = T1.UgpEntry  where  T0.DocEntry ='" & DocEntry & "' and t0.ItemCode ='" & ItemCode & "'"

                Case "T"
                    StrQuery = "select  T0. ItemCode, T0.UomCode, T0.UoMEntry, T1.UgpEntry , T1. InvntryUom ,T2.AltQty ,T2.BaseQty from WTR1 T0  left outer join OITM T1 on T1.ItemCode = T0.ItemCode " & _
                               "left outer join UGP1 T2 on T2.UomEntry = T0.UomEntry  and T2.UgpEntry = T1.UgpEntry  where  T0.DocEntry ='" & DocEntry & "' and t0.ItemCode ='" & ItemCode & "'"
                Case "R"
                    StrQuery = "select  T0. ItemCode, T0.UomCode, T0.UoMEntry, T1.UgpEntry , T1. InvntryUom ,T2.AltQty ,T2.BaseQty from IGN1 T0  left outer join OITM T1 on T1.ItemCode = T0.ItemCode " & _
                               "left outer join UGP1 T2 on T2.UomEntry = T0.UomEntry  and T2.UgpEntry = T1.UgpEntry  where  T0.DocEntry ='" & DocEntry & "' and t0.ItemCode ='" & ItemCode & "'"
            End Select
        End If
        objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objRecordSet.DoQuery(StrQuery)

        If objRecordSet.EoF Then Return Quantity
        Dim AltQty As Double = objRecordSet.Fields.Item("AltQty").Value
        Dim BaseQty As Double = objRecordSet.Fields.Item("BaseQty").Value

        convertedQty = BaseQty / AltQty * Quantity


        Return convertedQty
    End Function
    Private Sub getUOMValue(ByVal ItemCode As String, ByVal Type As String) ' not required

        Dim UOMCode As String
        Dim UgpEntry As Integer
        Select Case Type
            Case "G"
                strQuery = "select T1.UomCode,T0.UgpEntry from OITM T0 join PDN1 T1 on T0.ItemCode= T1.ItemCode where T1.itemCode='" & ItemCode & "'"
             
            Case "P"
                strQuery = "select T1.UomCode,T0.UgpEntry from OITM T0 join IGN1 T1 on T0.ItemCode= T1.ItemCode where T1.itemCode='" & ItemCode & "'"
               
            Case "T"
                strQuery = "select T1.UomCode,T0.UgpEntry from OITM T0 join WTR1 T1 on T0.ItemCode= T1.ItemCode where T1.itemCode='" & ItemCode & "'"
            Case "R"
                strQuery = "select T1.UomCode,T0.UgpEntry from OITM T0 join IGN1 T1 on T0.ItemCode= T1.ItemCode where T1.itemCode='" & ItemCode & "'"
        End Select
        objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objRecordSet.DoQuery(strQuery)

        UOMCode = objRecordSet.Fields.Item("UomCode").Value
        UgpEntry = objRecordSet.Fields.Item("UgpEntry").Value

    End Sub

    Public Sub getItemRemovalNotify(ByVal Type As String)
        Dim strQuery As String = ""
        Dim objItemres As SAPbobsCOM.Recordset
        If objAddOn.HANA Then
            Select Case Type
                Case "G"
                    strQuery = "select COUNT(*) From  PDN1 T0  join OITM T1 on T0.""ItemCode""=T1.""ItemCode"" where ifnull(T1.""U_InspReq"",'N')='N'AND T0.""DocEntry""='" & objForm.Items.Item("13B").Specific.String & "'"


                Case "P"
                    strQuery = "select COUNT(*) From  IGN1 T0  join OITM T1 on T0.""ItemCode""=T1.""ItemCode"" where ifnull(T1.""U_InspReq"",'N')='N'AND T0.""DocEntry""='" & objForm.Items.Item("51").Specific.String & "'"

                Case "T"
                    strQuery = "select COUNT(*) From  WTR1 T0  join OITM T1 on T0.""ItemCode""=T1.""ItemCode"" where ifnull(T1.""U_InspReq"",'N')='N'AND T0.""DocEntry""='" & objForm.Items.Item("23B").Specific.String & "'"

                Case "R"
                    strQuery = "select COUNT(*) From  IGN1 T0  join OITM T1 on T0.""ItemCode""=T1.""ItemCode"" where ifnull(T1.""U_InspReq"",'N')='N'AND T0.""DocEntry""='" & objForm.Items.Item("51").Specific.String & "'"
            End Select
        Else
            Select Case Type
                Case "G"
                    strQuery = "select COUNT(*) From  PDN1 T0  join OITM T1 on T0.itemcode=T1.itemcode where isnull(T1.U_inspreq,'N')='N'AND T0.DocEntry='" & objForm.Items.Item("13B").Specific.String & "'"

                Case "P"
                    strQuery = "select COUNT(*) From  IGN1 T0  join OITM T1 on T0.itemcode=T1.itemcode where isnull(T1.U_inspreq,'N')='N'AND T0.DocEntry='" & objForm.Items.Item("51").Specific.String & "'"


                Case "T"
                    strQuery = "select COUNT(*) From  WTR1 T0  join OITM T1 on T0.itemcode=T1.itemcode where isnull(T1.U_inspreq,'N')='N'AND T0.DocEntry='" & objForm.Items.Item("23B").Specific.String & "'"
                Case "R"
                    strQuery = "select COUNT(*) From  IGN1 T0  join OITM T1 on T0.itemcode=T1.itemcode where isnull(T1.U_inspreq,'N')='N'AND T0.DocEntry='" & objForm.Items.Item("51").Specific.String & "'"

            End Select
        End If
        objItemres = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objItemres.DoQuery(strQuery)
        If CInt(objItemres.Fields.Item(0).Value) > 0 Then
            objAddOn.objApplication.SetStatusBarMessage(CStr(objItemres.Fields.Item(0).Value) & " items Removed... Inspection Not Required", SAPbouiCOM.BoMessageTime.bmt_Short, False)
        End If

    End Sub
    
    Function validate(ByVal FormUID As String) As Boolean
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("20").Specific

        For i As Integer = 1 To objMatrix.VisualRowCount
            objCombo = objMatrix.Columns.Item("7AA").Cells.Item(i).Specific
            If objMatrix.Columns.Item("1").Cells.Item(i).Specific.string <> "" And CInt(objMatrix.Columns.Item("7").Cells.Item(i).Specific.string) > 0 And objCombo.Selected.Value = "-" Then
                objAddOn.objApplication.SetStatusBarMessage("Rejection Location has to be entered... Line: " & i, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                Return False
            End If
        Next
        'For intloop = 1 To objMatrix.RowCount
        '    If Not InspectedQtyCorrect(FormUID, intloop) Then
        '        Return False
        '    End If
        '    If Not RejDetailAvailable(FormUID) Then
        '        Return False
        '    End If
        'Next
        Return True
    End Function
    Function RejDetailAvailable(ByVal FormUID As String) As Boolean
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("20").Specific
        For intloop = 1 To objMatrix.RowCount
            If CInt(objMatrix.Columns.Item("7").Cells.Item(intloop).Specific.string) > 0 And objMatrix.Columns.Item("7_1").Cells.Item(intloop).Specific.string = "" Then
                objAddOn.objApplication.SetStatusBarMessage("Rejection Detail has to be entered")
                Return False
            End If
        Next
        Return True
    End Function
    Function getDetailsQuery(ByVal DocumentEntry As String, ByVal Type As String) As String
        ' should return one of below query
        Dim strSQL1 As String = ""
        Select Case Type
            Case "G"
                If objAddOn.HANA Then
                    strSQL1 = "SELECT ROW_NUMBER() OVER () AS ""LineId"",T2.""LineNum"", T2.""ItemCode"", T2.""Dscription"", T2.""Price"", T3.""InvntryUom"", T4.""BaseQty""/T4.""AltQty"" * Sum(T2.""Quantity"") AS ""TotQty"", (SELECT IFNULL(SUM(T1.""U_QtyInsp""), 0) FROM ""@MIPLQC1"" T1 " &
                        " INNER JOIN ""@MIPLQC"" T0 ON  T0.""DocEntry""=T1.""DocEntry"" AND T0.""U_GRNEntry"" = cast( T2.""DocEntry"" as varchar) AND T1.""U_ItemCode"" = T2.""ItemCode"" AND T1.""U_BaseLinNum"" =cast(T2.""LineNum"" as varchar)) AS ""InspQty"",  T4.""BaseQty""/T4.""AltQty"" * Sum(T2.""Quantity"") - (SELECT IFNULL(SUM(T1.""U_QtyInsp""), 0) FROM ""@MIPLQC1"" T1 " &
                         " INNER JOIN ""@MIPLQC"" T0 ON T0.""DocEntry""=T1.""DocEntry"" AND  T0.""U_GRNEntry"" = cast( T2.""DocEntry"" as varchar) AND T1.""U_ItemCode"" = T2.""ItemCode"" AND T1.""U_BaseLinNum"" =cast(T2.""LineNum"" as varchar)) AS ""PendQty"" FROM PDN1 T2 " &
                        " INNER JOIN ""OITM"" T3 ON T3.""ItemCode""=T2.""ItemCode"" left outer join UGP1 T4 on T4.""UomEntry"" = T2.""UomEntry""  and T4.""UgpEntry"" = T3.""UgpEntry"" WHERE T2.""DocEntry"" = '" & DocumentEntry & "' AND IFNULL(T3.""U_InspReq"",'') = 'Y' GROUP BY T2.""DocEntry"", T2.""ItemCode"", T2.""LineNum"", T2.""ItemCode"", T2.""Dscription"",T2.""Price"", T3.""InvntryUom"",T4.""BaseQty"",T4.""AltQty"";"

                Else
                    ' strSQL = "Select LineNum, ItemCode, Dscription,Quantity from PDN1 where DocEntry= '" & objForm.Items.Item("13B").Specific.string & "'"

                    strSQL1 = "select T2.LineNum, T2.ItemCode,T2. Dscription,T2.Price, T3.InvntryUom,T4.BaseQty/T4.AltQty * Sum(T2.Quantity) TotQty , (select isnull(sum(T1.U_QtyInsp),0) from [@MIPLQC1] T1 join [@MIPLQC] T0 on T0.DocEntry=T1.DocEntry and T0. U_GRNEntry = T2.DocEntry and T1.U_ItemCode = T2.ItemCode and T1.U_BaseLinNum =T2.LineNum) InspQty ,T4.BaseQty/T4.AltQty * Sum(T2.Quantity) - (select isnull(sum(T1.U_QtyInsp),0)   from [@MIPLQC1] T1 join [@MIPLQC] T0 on  T0.DocEntry=T1.DocEntry and T0. U_GRNEntry = T2.DocEntry and T1.U_ItemCode = T2.ItemCode and T1.U_BaseLinNum =T2.LineNum) PendQty" & _
        " from PDN1 T2 inner join OITM T3 on T3.ItemCode= T2.ItemCode left outer join UGP1 T4 on T4.UomEntry = T2.UomEntry  and T4.UgpEntry = T3.UgpEntry where T2.DocEntry ='" & DocumentEntry & "' AND isnull(T3.U_InspReq,'')='Y' group by T2.DocEntry ,T2.ItemCode ,T2.LineNum, T2.ItemCode,T2. Dscription, T2.price,T3.InvntryUom,T4.BaseQty,T4.AltQty"
                End If
            Case "P"
                If objAddOn.HANA Then
                    strSQL1 = " SELECT ROW_NUMBER() OVER () AS ""LineId"",T2.""LineNum"", T2.""ItemCode"", T2.""Dscription"", T2.""Price"", T3.""InvntryUom"",  T4.""BaseQty""/T4.""AltQty"" * Sum(T2.""Quantity"") AS ""TotQty"", (SELECT IFNULL(SUM(T1.""U_QtyInsp""), 0) " &
                        " FROM ""@MIPLQC1"" T1 INNER JOIN ""@MIPLQC"" T0 ON T0.""DocEntry"" = T1.""DocEntry"" AND T0.""U_GREntry"" = cast( T2.""DocEntry"" as varchar) AND T1.""U_ItemCode"" = T2.""ItemCode"" " &
                        " AND T1.""U_BaseLinNum"" = cast(T2.""LineNum"" as varchar)) AS ""InspQty"",  T4.""BaseQty""/T4.""AltQty"" * Sum(T2.""Quantity"") - (SELECT IFNULL(SUM(T1.""U_QtyInsp""), 0) " &
                        " FROM ""@MIPLQC1"" T1 INNER JOIN ""@MIPLQC"" T0 ON T0.""DocEntry"" = T1.""DocEntry"" AND T0.""U_GREntry"" = cast( T2.""DocEntry"" as varchar) AND T1.""U_ItemCode"" = T2.""ItemCode"" " &
                        " AND T1.""U_BaseLinNum"" = cast(T2.""LineNum"" as varchar)) AS ""PendQty"" FROM IGN1 T2  INNER JOIN ""OITM"" T3 ON T3.""ItemCode""=T2.""ItemCode"" left outer join UGP1 T4 on T4.""UomEntry"" = T2.""UomEntry""  and T4.""UgpEntry"" = T3.""UgpEntry"" WHERE T2.""DocEntry"" = '" & DocumentEntry & "' AND T2.""BaseType"" = '202' AND IFNULL(T3.""U_InspReq"",'') = 'Y' GROUP BY T2.""DocEntry"", T2.""ItemCode"", T2.""LineNum"", T2.""ItemCode"", T2.""Dscription"", T2.""Price"", T3.""InvntryUom"",T4.""BaseQty"",T4.""AltQty"";"

                Else
                    strSQL1 = "select T2.LineNum, T2.ItemCode,T2. Dscription,T2.Price, T3.InvntryUom, T4.BaseQty/T4.AltQty * Sum(T2.Quantity) TotQty , (select isnull(sum(T1.U_QtyInsp),0) from [@MIPLQC1] T1 " & _
            " join [@MIPLQC] T0 on T0.DocEntry=T1.DocEntry and T0. U_GREntry = T2.DocEntry and T1.U_ItemCode = T2.ItemCode and T1.U_BaseLinNum =T2.LineNum) InspQty, " & _
  " T4.BaseQty/T4.AltQty * Sum(T2.Quantity) - (select isnull(sum(T1.U_QtyInsp),0)   from [@MIPLQC1] T1 join [@MIPLQC] T0 on  T0.DocEntry=T1.DocEntry and " & _
   " T0. U_GREntry = T2.DocEntry and T1.U_ItemCode = T2.ItemCode and T1.U_BaseLinNum =T2.LineNum) PendQty " & _
         " from IGN1 T2  inner join OITM T3 on T3.ItemCode= T2.ItemCode left outer join UGP1 T4 on T4.UomEntry = T2.UomEntry  and T4.UgpEntry = T3.UgpEntry where T2.DocEntry ='" & DocumentEntry & "'  and T2.BaseType='202'  AND isnull(T3.U_InspReq,'')='Y' group by T2.DocEntry ,T2.ItemCode ,T2.LineNum, T2.ItemCode,T2. Dscription, T2.price, T3.InvntryUom ,T4.BaseQty,T4.AltQty"


                End If
            Case "T"
                If objAddOn.HANA Then
                    strSQL1 = "SELECT ROW_NUMBER() OVER () AS ""LineId"",T2.""LineNum"", T2.""ItemCode"", T2.""Dscription"", T2.""StockPrice"" AS ""Price"", T3.""InvntryUom"",  T4.""BaseQty""/T4.""AltQty"" * Sum(T2.""Quantity"") AS ""TotQty"", (SELECT IFNULL(SUM(T1.""U_QtyInsp""), 0) FROM ""@MIPLQC1"" T1 " &
                        " INNER JOIN ""@MIPLQC"" T0 ON T0.""DocEntry""=T1.""DocEntry"" AND  T0.""U_TransEntry"" = cast( T2.""DocEntry"" as varchar) AND T1.""U_ItemCode"" = T2.""ItemCode"" AND T1.""U_BaseLinNum"" =cast(T2.""LineNum"" as varchar)) AS ""InspQty"",  T4.""BaseQty""/T4.""AltQty"" * Sum(T2.""Quantity"") - (SELECT IFNULL(SUM(T1.""U_QtyInsp""), 0) FROM ""@MIPLQC1"" T1 " &
                        " INNER JOIN ""@MIPLQC"" T0 ON T0.""DocEntry""=T1.""DocEntry"" AND  T0.""U_TransEntry"" = cast( T2.""DocEntry"" as varchar) AND T1.""U_ItemCode"" = T2.""ItemCode"" AND T1.""U_BaseLinNum"" =cast(T2.""LineNum"" as varchar)) AS ""PendQty"" FROM WTR1 T2 " &
                        "  INNER JOIN ""OITM"" T3 ON T3.""ItemCode""=T2.""ItemCode"" left outer join UGP1 T4 on T4.""UomEntry"" = T2.""UomEntry""  and T4.""UgpEntry"" = T3.""UgpEntry"" WHERE T2.""DocEntry"" = '" & DocumentEntry & "'  AND IFNULL(T3.""U_InspReq"",'') = 'Y' GROUP BY T2.""DocEntry"", T2.""ItemCode"", T2.""LineNum"", T2.""ItemCode"", T2.""Dscription"", T2.""StockPrice"", T3.""InvntryUom"",T4.""BaseQty"",T4.""AltQty"";"
                Else
                    strSQL1 = "select T2.LineNum, T2.ItemCode,T2. Dscription,T2.StockPrice AS Price, T3.InvntryUom, T4.BaseQty/T4.AltQty * Sum(T2.Quantity) TotQty , (select isnull(sum(T1.U_QtyInsp),0) from [@MIPLQC1] T1 " & _
               " join [@MIPLQC] T0 on  T0.DocEntry=T1.DocEntry and T0. U_TransEntry = T2.DocEntry and T1.U_ItemCode = T2.ItemCode and T1.U_BaseLinNum =T2.LineNum) InspQty, T4.BaseQty/T4.AltQty * Sum(T2.Quantity) - (select isnull(sum(T1.U_QtyInsp),0)   from [@MIPLQC1] T1 join [@MIPLQC] T0 on  T0.DocEntry=T1.DocEntry and T0. U_TransEntry = T2.DocEntry and T1.U_ItemCode = T2.ItemCode and T1.U_BaseLinNum =T2.LineNum) PendQty " & _
              " from WTR1 T2  inner join OITM T3 on T3.ItemCode= T2.ItemCode left outer join UGP1 T4 on T4.UomEntry = T2.UomEntry  and T4.UgpEntry = T3.UgpEntry where T2.DocEntry ='" & DocumentEntry & "'  AND isnull(T3.U_InspReq,'')='Y' group by T2.DocEntry ,T2.ItemCode ,T2.LineNum, T2.ItemCode,T2. Dscription, T2.StockPrice, T3.InvntryUom ,T4.BaseQty,T4.AltQty"

                End If
            Case "R"
                If objAddOn.HANA Then
                    strSQL1 = " SELECT ROW_NUMBER() OVER () AS ""LineId"",T2.""LineNum"", T2.""ItemCode"", T2.""Dscription"", T2.""Price"", T3.""InvntryUom"",  T4.""BaseQty""/T4.""AltQty"" * Sum(T2.""Quantity"") AS ""TotQty"", (SELECT IFNULL(SUM(T1.""U_QtyInsp""), 0) " &
                        " FROM ""@MIPLQC1"" T1 INNER JOIN ""@MIPLQC"" T0 ON T0.""DocEntry"" = T1.""DocEntry"" AND T0.""U_GREntry"" = cast( T2.""DocEntry"" as varchar) AND T1.""U_ItemCode"" = T2.""ItemCode"" " &
                        " AND T1.""U_BaseLinNum"" =cast( T2.""LineNum"" as varchar)) AS ""InspQty"",  T4.""BaseQty""/T4.""AltQty"" * Sum(T2.""Quantity"") - (SELECT IFNULL(SUM(T1.""U_QtyInsp""), 0) " &
                        " FROM ""@MIPLQC1"" T1 INNER JOIN ""@MIPLQC"" T0 ON T0.""DocEntry"" = T1.""DocEntry"" AND T0.""U_GREntry"" = cast( T2.""DocEntry"" as varchar) AND T1.""U_ItemCode"" = T2.""ItemCode"" " &
                        " AND T1.""U_BaseLinNum"" =cast( T2.""LineNum"" as varchar)) AS ""PendQty"" FROM IGN1 T2  INNER JOIN ""OITM"" T3 ON T3.""ItemCode""=T2.""ItemCode"" left outer join UGP1 T4 on T4.""UomEntry"" = T2.""UomEntry""  and T4.""UgpEntry"" = T3.""UgpEntry"" WHERE T2.""DocEntry"" = '" & DocumentEntry & "' AND T2.""BaseType"" <>'202'  AND IFNULL(T3.""U_InspReq"",'') = 'Y' GROUP BY T2.""DocEntry"", T2.""ItemCode"", T2.""LineNum"", T2.""ItemCode"", T2.""Dscription"", T2.""Price"", T3.""InvntryUom"",T4.""BaseQty"",T4.""AltQty"";"

                Else
                    strSQL1 = "select T2.LineNum, T2.ItemCode,T2. Dscription,T2.Price, T3.InvntryUom, T4.BaseQty/T4.AltQty * Sum(T2.Quantity) TotQty , (select isnull(sum(T1.U_QtyInsp),0) from [@MIPLQC1] T1 " & _
            " join [@MIPLQC] T0 on T0.DocEntry=T1.DocEntry and T0. U_GREntry = T2.DocEntry and T1.U_ItemCode = T2.ItemCode and T1.U_BaseLinNum =T2.LineNum) InspQty, " & _
  " T4.BaseQty/T4.AltQty * Sum(T2.Quantity) - (select isnull(sum(T1.U_QtyInsp),0)   from [@MIPLQC1] T1 join [@MIPLQC] T0 on  T0.DocEntry=T1.DocEntry and " & _
   " T0. U_GREntry = T2.DocEntry and T1.U_ItemCode = T2.ItemCode and T1.U_BaseLinNum =T2.LineNum) PendQty " & _
         " from IGN1 T2  inner join OITM T3 on T3.ItemCode= T2.ItemCode left outer join UGP1 T4 on T4.UomEntry = T2.UomEntry  and T4.UgpEntry = T3.UgpEntry where T2.DocEntry ='" & DocumentEntry & "'  and T2.BaseType<>'202'  AND isnull(T3.U_InspReq,'')='Y' group by T2.DocEntry ,T2.ItemCode ,T2.LineNum, T2.ItemCode,T2. Dscription, T2.price, T3.InvntryUom ,T4.BaseQty,T4.AltQty"


                End If

        End Select
        Return strSQL1
    End Function
    Function getDocumentEntry(ByVal FormUID As String, ByVal Type As String) As String
        Dim DocumentEntry As String = ""
        Select Case Type
            Case "G"

                If objAddOn.HANA Then
                    strSQL = "Select Top 1 T0.""DocEntry"", T1.""WhsCode"", T0.""CardName"" from OPDN T0 join PDN1 T1 on T0.""DocEntry"" = T1.""DocEntry"" where T0.""InvntSttus""='O' AND T0.""DocNum"" ='" & objForm.Items.Item("13").Specific.string & "' order by ""DocEntry"" Desc"
                Else
                    strSQL = "Select Top 1 T0.DocEntry, T1.WhsCode, T0.CardName from OPDN T0 join PDN1 T1 on T0.DocEntry = T1.DocEntry where T0.InvntSttus='O' AND T0.DocNum ='" & objForm.Items.Item("13").Specific.string & "' order by DocEntry Desc"
                End If

                objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRecordSet.DoQuery(strSQL)
                If Not objRecordSet.EoF Then
                    DocumentEntry = CStr(objRecordSet.Fields.Item("DocEntry").Value)
                    objForm.Items.Item("13B").Specific.string = DocumentEntry
                    objForm.Items.Item("25").Specific.string = CStr(objRecordSet.Fields.Item("WhsCode").Value)
                    objForm.Items.Item("27").Specific.string = CStr(objRecordSet.Fields.Item("CardName").Value)
                End If
                objRecordSet = Nothing
            Case "P"
                If objAddOn.HANA Then
                    strSQL = "Select Top 1 ""DocEntry"", ""Warehouse"" from OWOR where ""DocNum"" ='" & objForm.Items.Item("15").Specific.string & "' order by ""DocEntry"" Desc"
                Else
                    strSQL = "Select Top 1 DocEntry, Warehouse from OWOR where DocNum ='" & objForm.Items.Item("15").Specific.string & "' order by DocEntry Desc"
                End If

                objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRecordSet.DoQuery(strSQL)
                If Not objRecordSet.EoF Then
                    DocumentEntry = CStr(objRecordSet.Fields.Item("DocEntry").Value)
                    objForm.Items.Item("15B").Specific.string = DocumentEntry
                    objForm.Items.Item("25").Specific.string = CStr(objRecordSet.Fields.Item("Warehouse").Value)
                End If
                objRecordSet = Nothing


            Case "T"
                If objAddOn.HANA Then
                    strSQL = "Select Top 1 T0.""DocEntry"", T1.""WhsCode"", T0.""CardName"" from OWTR T0 join WTR1 T1 on T0.""DocEntry"" = T1.""DocEntry"" where T0.""DocNum"" ='" & objForm.Items.Item("23").Specific.string & "' order by ""DocEntry"" Desc"
                Else
                    strSQL = "Select Top 1 T0.DocEntry, T1.WhsCode, T0.CardName from OWTR T0 join WTR1 T1 on T0.DocEntry = T1.DocEntry where T0.DocNum ='" & objForm.Items.Item("23").Specific.string & "' order by DocEntry Desc"
                End If

                objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRecordSet.DoQuery(strSQL)
                If Not objRecordSet.EoF Then
                    DocumentEntry = CStr(objRecordSet.Fields.Item("DocEntry").Value)
                    objForm.Items.Item("23B").Specific.string = DocumentEntry
                    objForm.Items.Item("25").Specific.string = CStr(objRecordSet.Fields.Item("WhsCode").Value)
                    objForm.Items.Item("27").Specific.string = CStr(objRecordSet.Fields.Item("CardName").Value)
                End If
                objRecordSet = Nothing
            Case "R"
                If objAddOn.HANA Then
                    strSQL = "Select Top 1 T0.""DocEntry"", T1.""WhsCode"", T0.""CardName"", T0.""DocNum"" from OIGN T0 join IGN1 T1 on T0.""DocEntry"" = T1.""DocEntry"" where T0.""InvntSttus""='O' AND T0.""DocEntry"" ='" & objForm.Items.Item("51").Specific.string & "' order by T0.""DocEntry"" Desc"
                Else
                    strSQL = "Select Top 1 T0.DocEntry, T1.WhsCode, T0.CardName, T0.DocNum from OIGN T0 join IGN1 T1 on T0.DocEntry = T1.DocEntry where T0.InvntSttus='O' AND T0.DocEntry ='" & objForm.Items.Item("51").Specific.string & "' order by T0.DocEntry Desc"
                End If

                objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRecordSet.DoQuery(strSQL)
                If Not objRecordSet.EoF Then
                    DocumentEntry = CStr(objRecordSet.Fields.Item("DocEntry").Value)
                    objForm.Items.Item("51B").Specific.string = objRecordSet.Fields.Item("DocNum").Value
                    objForm.Items.Item("25").Specific.string = CStr(objRecordSet.Fields.Item("WhsCode").Value)
                End If
                objRecordSet = Nothing

        End Select
        Return DocumentEntry
    End Function

    Sub DeleteRow()
        Try
            objMatrix = objForm.Items.Item("20").Specific
            objMatrix.FlushToDataSource()
            For i As Integer = 1 To objMatrix.VisualRowCount
                objMatrix.GetLineData(i)
                QCLine.Offset = i - 1
                QCLine.SetValue("LineId", QCLine.Offset, i)
                objMatrix.SetLineData(i)
                objMatrix.FlushToDataSource()
            Next
            QCLine.RemoveRecord(QCLine.Size - 1)
            objMatrix.LoadFromDataSource()

        Catch ex As Exception
            objAddOn.objApplication.StatusBar.SetText("DeleteRow  Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Public Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try

            objForm = objAddOn.objApplication.Forms.Item(objForm.UniqueID)
            objMatrix = objForm.Items.Item("20").Specific
            Dim ItemUID As String = ""
            Select Case pVal.MenuUID
                Case "1284"

                Case "1281" 'Find Mode
                    If pVal.BeforeAction = False Then
                        objForm.Items.Item("4").Enabled = True
                        objForm.Items.Item("6").Enabled = True
                        objForm.Items.Item("6E").Enabled = True
                        objForm.Items.Item("6C").Enabled = True
                        objForm.Items.Item("25").Enabled = True
                        objForm.Items.Item("27").Enabled = True
                        objForm.Items.Item("13B").Enabled = True
                        objForm.Items.Item("23B").Enabled = True
                        objForm.Items.Item("15B").Enabled = True
                        objForm.Items.Item("51B").Enabled = True
                    End If
                Case "1282"
                    If pVal.BeforeAction = False Then
                        objCombo = objForm.Items.Item("50").Specific
                        objCombo.Select("-1", SAPbouiCOM.BoSearchKey.psk_ByValue)
                        objCombo = objForm.Items.Item("10").Specific
                        objCombo.Select("-", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    End If

                Case "1293"  'delete Row
                    If ItemUID = "20" Then
                        'DeleteRow()
                    End If

            End Select

        Catch ex As Exception
            objAddOn.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
        End Try
    End Sub
    Sub RightClickEvent(ByRef EventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Try
            objForm = objAddOn.objApplication.Forms.Item(objForm.UniqueID)
            objMatrix = objForm.Items.Item("20").Specific
            If EventInfo.BeforeAction Then
                Select Case EventInfo.EventType
                    Case SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK
                        Select Case EventInfo.ItemUID
                            Case "20"
                                If EventInfo.ColUID = "0" And objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    objForm.EnableMenu("1293", True)
                                Else
                                    objForm.EnableMenu("1293", False)
                                End If
                        End Select
                End Select
            Else

            End If

            'If EventInfo.BeforeAction Then
            'Else
            '    Select Case EventInfo.ItemUID
            '        Case "20"
            '            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            '                objForm.EnableMenu("1293", True)
            '            End If

            '    End Select
            'End If

        Catch ex As Exception
            'objAddOn.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
        End Try
    End Sub
    Function StockPosting_GRN(ByVal FormUID As String) As Boolean
        Dim oStockTransfer As SAPbobsCOM.StockTransfer
        Dim BinWhse As String = ""
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("20").Specific
        If objMatrix.RowCount = 0 Then
            Return False
        End If
        Try

            Dim SumAccQty, SumRejQty, SumRewQty As Integer
            SumAccQty = SumRejQty = SumRewQty = 0

            For intloop = 1 To objMatrix.RowCount
                SumAccQty += objMatrix.Columns.Item("6").Cells.Item(intloop).Specific.string
                SumRejQty += objMatrix.Columns.Item("7").Cells.Item(intloop).Specific.string
                SumRewQty += objMatrix.Columns.Item("8").Cells.Item(intloop).Specific.string
                If BinWhse = "" Then
                    BinWhse = BinEnabled(objMatrix.Columns.Item("6A").Cells.Item(intloop).Specific.string)
                End If
            Next intloop


            InWhse = objForm.Items.Item("25").Specific.string
            If SumAccQty > 0 And BinWhse = "" Then
                oStockTransfer = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer)
                Dim DocNum = objForm.Items.Item("4").Specific.string
                oStockTransfer.Comments = "Accepted Stock Posted From QC DocNum -> " & DocNum
                oStockTransfer.Reference2 = DocNum
                oStockTransfer.DocDate = objAddOn.objGenFunc.GetDateTimeValue(objForm.Items.Item("6").Specific.String) '.ToString("yyyyMMdd") 'Date.Now
                oStockTransfer.FromWarehouse = InWhse
                Dim ItemCode = ""
                For intloop = 1 To objMatrix.RowCount
                    AccQty = IIf(objMatrix.Columns.Item("6").Cells.Item(intloop).Specific.string.Trim = "", 0, CInt(objMatrix.Columns.Item("6").Cells.Item(intloop).Specific.string))
                    If AccQty > 0 Then
                        oStockTransfer.ToWarehouse = objMatrix.Columns.Item("6A").Cells.Item(intloop).Specific.string
                        ItemCode = objMatrix.Columns.Item("1").Cells.Item(intloop).Specific.string

                        If objAddOn.HANA Then
                            strSQL = "SELECT IFNULL(""InvntItem"", 'N') AS ""InvtItem"" FROM OITM WHERE ""ItemCode"" = '" & ItemCode & "';"
                        Else
                            strSQL = "Select isnull(InvntItem,'N') InvtItem from OITM where ItemCode = '" & ItemCode & "'"
                        End If
                        Dim ChkInvt = objAddOn.objGenFunc.getSingleValue(strSQL)
                        If ChkInvt = "N" Then Return True
                        oStockTransfer.Lines.ItemCode = ItemCode
                        oStockTransfer.Lines.Quantity = AccQty
                        oStockTransfer.Lines.FromWarehouseCode = InWhse
                        oStockTransfer.Lines.WarehouseCode = objMatrix.Columns.Item("6A").Cells.Item(intloop).Specific.string
                        oStockTransfer.Lines.BinAllocations.BinAbsEntry = 3
                        oStockTransfer.Lines.BinAllocations.Quantity = AccQty
                        oStockTransfer.Lines.BinAllocations.SerialAndBatchNumbersBaseLine = 0
                        oStockTransfer.Lines.BinAllocations.Add()
                        If objAddOn.HANA Then
                            strSQL = "Select ""ManBtchNum""  from OITM Where ""ItemCode""='" & ItemCode & "';"
                        Else
                            strSQL = "Select ManBtchNum  from OITM Where ItemCode='" & ItemCode & "'"
                        End If

                        Dim ManBtchNum = objAddOn.objGenFunc.getSingleValue(strSQL)

                        If ManBtchNum.Trim <> "N" Then

                            If objAddOn.HANA Then
                                strSQL = "select * from OBTQ where ""ItemCode"" ='" & ItemCode & "' and ""WhsCode"" ='" & InWhse & "' And ""Quantity"" > 0 order by ""SysNumber"""
                            Else
                                strSQL = "select * from OBTQ where ItemCode ='" & ItemCode & "' and WhsCode ='" & InWhse & "' And Quantity > 0 order by SysNumber"

                            End If
                            Dim objRecordSet As SAPbobsCOM.Recordset = objAddOn.objGenFunc.DoQuery(strSQL)

                            Dim count As Double = CDbl(AccQty)

                            For k As Integer = 0 To objRecordSet.RecordCount - 1
                                If objAddOn.HANA Then
                                    strSQL = "select ""DistNumber"" from OBTN where ""ItemCode"" ='" & ItemCode & "' and ""SysNumber"" ='" & objRecordSet.Fields.Item("SysNumber").Value & "'"
                                Else
                                    strSQL = "select DistNumber from OBTN where ItemCode ='" & ItemCode & "' and SysNumber ='" & objRecordSet.Fields.Item("SysNumber").Value & "'"

                                End If
                                oStockTransfer.Lines.BatchNumbers.BatchNumber = objAddOn.objGenFunc.getSingleValue(strSQL)
                                If CDbl(objRecordSet.Fields.Item("Quantity").Value) >= count Then
                                    ' Dim ss = objRecordSet.Fields.Item("SysNumber").Value
                                    oStockTransfer.Lines.BatchNumbers.Quantity = count

                                    oStockTransfer.Lines.BatchNumbers.Add()
                                    Exit For
                                Else

                                    oStockTransfer.Lines.BatchNumbers.Quantity = objRecordSet.Fields.Item("Quantity").Value
                                    oStockTransfer.Lines.BatchNumbers.Add()
                                    count = count - CDbl(objRecordSet.Fields.Item("Quantity").Value)

                                End If
                                objRecordSet.MoveNext()
                            Next

                        End If

                        oStockTransfer.Lines.Add()
                    End If ' AccQty>0
                Next intloop ' accepted lines loop end
                Dim ErrCode = oStockTransfer.Add()
                If ErrCode <> 0 Then
                    objAddOn.objApplication.SetStatusBarMessage(" QC Acc Qty Posting Error : " & objAddOn.objCompany.GetLastErrorDescription)

                    Return False
                Else
                    QCHeader.SetValue("U_AccStk", 0, objAddOn.objCompany.GetNewObjectKey)
                    objAddOn.objApplication.SetStatusBarMessage("QC Accepted Quantity Posted", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                End If
            End If
            '--------------------------------- Rejected Qty Stock Transfer------------------------------------------

            If SumRejQty > 0 Then
                oStockTransfer = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer)

                Dim DocNum = objForm.Items.Item("4").Specific.string
                oStockTransfer.Comments = "Rejected Stock Posted From QC DocNum -> " & DocNum
                oStockTransfer.Reference2 = DocNum
                oStockTransfer.DocDate = objAddOn.objGenFunc.GetDateTimeValue(objForm.Items.Item("6").Specific.String) 'Date.Now
                oStockTransfer.FromWarehouse = InWhse


                Dim ItemCode = ""
                Dim ToRejWorkCenter = ""
                Dim ToRewWorkCenter = ""
                For intloop = 1 To objMatrix.RowCount
                    RejQty = IIf(objMatrix.Columns.Item("7").Cells.Item(intloop).Specific.string.Trim = "", 0, CInt(objMatrix.Columns.Item("7").Cells.Item(intloop).Specific.string))
                    If RejQty > 0 Then
                        oStockTransfer.ToWarehouse = objMatrix.Columns.Item("7A").Cells.Item(intloop).Specific.string
                        ItemCode = objMatrix.Columns.Item("1").Cells.Item(intloop).Specific.string
                        'Dim oFlag As Boolean = False
                        If objAddOn.HANA Then
                            strSQL = "SELECT IFNULL(""InvntItem"", 'N') AS ""InvtItem"" FROM OITM WHERE ""ItemCode"" = '" & ItemCode & "';"
                        Else
                            strSQL = "Select isnull(InvntItem,'N') InvtItem from OITM where ItemCode = '" & ItemCode & "'"
                        End If
                        Dim ChkInvt = objAddOn.objGenFunc.getSingleValue(strSQL)
                        If ChkInvt = "N" Then Return True
                        oStockTransfer.Lines.ItemCode = ItemCode
                        oStockTransfer.Lines.Quantity = RejQty
                        oStockTransfer.Lines.FromWarehouseCode = InWhse
                        oStockTransfer.Lines.WarehouseCode = objMatrix.Columns.Item("7A").Cells.Item(intloop).Specific.string

                        'Batch Number Allocation
                        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                        If objAddOn.HANA Then
                            strSQL = "Select ""ManBtchNum""  from OITM Where ""ItemCode""='" & ItemCode & "';"
                        Else
                            strSQL = "Select ManBtchNum  from OITM Where ItemCode='" & ItemCode & "'"
                        End If
                        Dim ManBtchNum = objAddOn.objGenFunc.getSingleValue(strSQL)
                        If ManBtchNum.Trim <> "N" Then
                            If objAddOn.HANA Then
                                strSQL = "select * from OBTQ where ""ItemCode"" ='" & ItemCode & "' and ""WhsCode"" ='" & InWhse & "' And ""Quantity"" > 0 order by ""SysNumber"""
                            Else
                                strSQL = "select * from OBTQ where ItemCode ='" & ItemCode & "' and WhsCode ='" & InWhse & "' And Quantity > 0 order by SysNumber"

                            End If
                            Dim objRecordSet As SAPbobsCOM.Recordset = objAddOn.objGenFunc.DoQuery(strSQL)

                            Dim count As Double = CDbl(RejQty)

                            For k As Integer = 0 To objRecordSet.RecordCount - 1
                                If objAddOn.HANA Then
                                    strSQL = "select ""DistNumber"" from OBTN where ""ItemCode"" ='" & ItemCode & "' and ""SysNumber"" ='" & objRecordSet.Fields.Item("SysNumber").Value & "'"
                                Else
                                    strSQL = "select DistNumber from OBTN where ItemCode ='" & ItemCode & "' and SysNumber ='" & objRecordSet.Fields.Item("SysNumber").Value & "'"

                                End If
                                oStockTransfer.Lines.BatchNumbers.BatchNumber = objAddOn.objGenFunc.getSingleValue(strSQL)
                                If CDbl(objRecordSet.Fields.Item("Quantity").Value) >= count Then
                                    ' Dim ss = objRecordSet.Fields.Item("SysNumber").Value
                                    oStockTransfer.Lines.BatchNumbers.Quantity = count
                                    oStockTransfer.Lines.BatchNumbers.Add()
                                    Exit For
                                Else

                                    oStockTransfer.Lines.BatchNumbers.Quantity = objRecordSet.Fields.Item("Quantity").Value
                                    oStockTransfer.Lines.BatchNumbers.Add()
                                    count = count - CDbl(objRecordSet.Fields.Item("Quantity").Value)

                                End If
                                objRecordSet.MoveNext()
                            Next

                        End If


                        oStockTransfer.Lines.Add()
                    End If ' RejQty>0
                Next intloop ' Rejected lines loop end
                Dim ErrCode = oStockTransfer.Add()
                If ErrCode <> 0 Then
                    objAddOn.objApplication.SetStatusBarMessage(" QC Rej Qty Posting Error : " & objAddOn.objCompany.GetLastErrorDescription)
                    Return False
                Else
                    QCHeader.SetValue("U_RejStk", 0, objAddOn.objCompany.GetNewObjectKey)
                    objAddOn.objApplication.SetStatusBarMessage(" QC Rejected Qty Posted ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                End If
            End If

            '------------------------Rework Qty Stock Transfer ------------------------------------------
            If SumRewQty > 0 Then
                oStockTransfer = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer)

                Dim DocNum = objForm.Items.Item("4").Specific.string
                oStockTransfer.Comments = "Rework stock Posted From QC DocNum -> " & DocNum
                oStockTransfer.Reference2 = DocNum
                oStockTransfer.DocDate = objAddOn.objGenFunc.GetDateTimeValue(objForm.Items.Item("6").Specific.String) ' Date.Now
                oStockTransfer.FromWarehouse = InWhse


                Dim ItemCode = ""
                Dim ToRejWorkCenter = ""
                Dim ToRewWorkCenter = ""
                For intloop = 1 To objMatrix.RowCount
                    RewQty = IIf(objMatrix.Columns.Item("8").Cells.Item(intloop).Specific.string.Trim = "", 0, CInt(objMatrix.Columns.Item("8").Cells.Item(intloop).Specific.string))
                    If RewQty > 0 Then
                        oStockTransfer.ToWarehouse = objMatrix.Columns.Item("8A").Cells.Item(intloop).Specific.string
                        ItemCode = objMatrix.Columns.Item("1").Cells.Item(intloop).Specific.string
                        'Dim oFlag As Boolean = False
                        If objAddOn.HANA Then
                            strSQL = "SELECT IFNULL(""InvntItem"", 'N') AS ""InvtItem"" FROM OITM WHERE ""ItemCode"" = '" & ItemCode & "';"
                        Else
                            strSQL = "Select isnull(InvntItem,'N') InvtItem from OITM where ItemCode = '" & ItemCode & "'"
                        End If
                        Dim ChkInvt = objAddOn.objGenFunc.getSingleValue(strSQL)
                        If ChkInvt = "N" Then Return True
                        oStockTransfer.Lines.ItemCode = ItemCode
                        oStockTransfer.Lines.Quantity = RewQty
                        oStockTransfer.Lines.FromWarehouseCode = InWhse
                        oStockTransfer.Lines.WarehouseCode = objMatrix.Columns.Item("8A").Cells.Item(intloop).Specific.string


                        'Batch Number Allocation
                        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                        If objAddOn.HANA Then
                            strSQL = "Select ""ManBtchNum""  from OITM Where ""ItemCode""='" & ItemCode & "';"
                        Else
                            strSQL = "Select ManBtchNum  from OITM Where ItemCode='" & ItemCode & "'"
                        End If
                        Dim ManBtchNum = objAddOn.objGenFunc.getSingleValue(strSQL)
                        If ManBtchNum.Trim <> "N" Then
                            If objAddOn.HANA Then
                                strSQL = "select * from OBTQ where ""ItemCode"" ='" & ItemCode & "' and ""WhsCode"" ='" & InWhse & "' And ""Quantity"" > 0 order by ""SysNumber"""
                            Else
                                strSQL = "select * from OBTQ where ItemCode ='" & ItemCode & "' and WhsCode ='" & InWhse & "' And Quantity > 0 order by SysNumber"

                            End If
                            Dim objRecordSet As SAPbobsCOM.Recordset = objAddOn.objGenFunc.DoQuery(strSQL)

                            Dim count As Double = CDbl(RewQty)

                            For k As Integer = 0 To objRecordSet.RecordCount - 1
                                If objAddOn.HANA Then
                                    strSQL = "select ""DistNumber"" from OBTN where ""ItemCode"" ='" & ItemCode & "' and ""SysNumber"" ='" & objRecordSet.Fields.Item("SysNumber").Value & "'"
                                Else
                                    strSQL = "select DistNumber from OBTN where ItemCode ='" & ItemCode & "' and SysNumber ='" & objRecordSet.Fields.Item("SysNumber").Value & "'"

                                End If
                                oStockTransfer.Lines.BatchNumbers.BatchNumber = objAddOn.objGenFunc.getSingleValue(strSQL)
                                If CDbl(objRecordSet.Fields.Item("Quantity").Value) >= count Then
                                    ' Dim ss = objRecordSet.Fields.Item("SysNumber").Value
                                    oStockTransfer.Lines.BatchNumbers.Quantity = count
                                    oStockTransfer.Lines.BatchNumbers.Add()
                                    Exit For
                                Else

                                    oStockTransfer.Lines.BatchNumbers.Quantity = objRecordSet.Fields.Item("Quantity").Value
                                    oStockTransfer.Lines.BatchNumbers.Add()
                                    count = count - CDbl(objRecordSet.Fields.Item("Quantity").Value)

                                End If
                                objRecordSet.MoveNext()
                            Next
                        End If
                        oStockTransfer.Lines.Add()
                    End If ' RewQty>0
                Next intloop ' Rework lines loop end
                Dim ErrCode = oStockTransfer.Add()
                If ErrCode <> 0 Then
                    objAddOn.objApplication.SetStatusBarMessage(" QC Rework Qty Posting Error : " & objAddOn.objCompany.GetLastErrorDescription)
                    Return False
                Else
                    QCHeader.SetValue("U_RewStk", 0, objAddOn.objCompany.GetNewObjectKey)
                    objAddOn.objApplication.SetStatusBarMessage("QC Rework Qty Posted", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                End If
            End If

            '----------------------------------End of Rework quantity------------------
            QCHeader.SetValue("U_StkPost", 0, "Y")
            Return True

        Catch ex As Exception
            objAddOn.objApplication.SetStatusBarMessage(" GRN Stock Posting Method Failed " & ex.Message)
            Return False
        End Try
    End Function
    Sub StockTransfer_BinLocation(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("20").Specific
        Dim SumAccQty, SumRejQty, SumRewQty As Integer
        Dim BinWhse As String = ""
        Dim QCEntry As String = ""
        If objForm.Items.Item("34").Specific.String = "" Then
            QCEntry = getQCEntry(FormUID)
            If QCEntry = "" Then
                SumAccQty = SumRejQty = SumRewQty = 0

                For intloop = 1 To objMatrix.RowCount
                    SumAccQty += objMatrix.Columns.Item("6").Cells.Item(intloop).Specific.string
                    SumRejQty += objMatrix.Columns.Item("7").Cells.Item(intloop).Specific.string
                    SumRewQty += objMatrix.Columns.Item("8").Cells.Item(intloop).Specific.string
                    If BinWhse = "" Then
                        BinWhse = BinEnabled(objMatrix.Columns.Item("6A").Cells.Item(intloop).Specific.string)
                    End If
                Next intloop

                If (SumAccQty > 0 Or SumRejQty > 0 Or SumRewQty > 0) And objForm.Items.Item("34").Specific.String = "" Then
                    objAddOn.objApplication.Menus.Item("3080").Activate()
                End If
            Else
                objForm.Items.Item("32").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                objForm.Items.Item("34").Specific.String = QCEntry
            End If
        End If

    End Sub
    Function BinEnabled(ByVal Whse As String) As String
        If objAddOn.HANA Then
            strSQL = "SELECT IFNULL(""BinActivat"", 'N') AS ""BinActivat"" FROM OWHS WHERE ""WhsCode"" = '" & Whse & "';"
        Else
            strSQL = "Select isnull(BinActivat,'N') BinActivat from OWHS where WhsCode = '" & Whse & "'"
        End If
        Dim ChkInvt = objAddOn.objGenFunc.getSingleValue(strSQL)
        If ChkInvt = "Y" Then Return Whse
        Return ""

    End Function
    Function getQCEntry(ByVal FormUID As String) As String
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        QCHeader = objForm.DataSources.DBDataSources.Item("@MIPLQC")
        If objAddOn.HANA Then
            strSQL = "SELECT Top 1 ""DocEntry"" FROM OWTR WHERE ""U_QCEntry"" = '" & objForm.Items.Item("6E").Specific.string & "' ORDER BY ""DocEntry"" DESC;"
        Else
            strSQL = "Select top 1 DocEntry  from OWTR where U_QCEntry = '" & objForm.Items.Item("6E").Specific.string & "' order by DocEntry DESC"
        End If
        Return objAddOn.objGenFunc.getSingleValue(strSQL)
    End Function
    Private Sub BatchUpdate()
        Dim objDoc As SAPbobsCOM.Documents
        objDoc = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes)
        Dim batch As SAPbobsCOM.InventoryPostingBatchNumber





        '    '' MsgBox(objDoc.Lines.ItemCode)
        If objDoc.GetByKey(27) Then

            objDoc.Lines.SetCurrentLine(0)

            objDoc.Lines.ItemCode = objDoc.Lines.ItemCode
            objDoc.Lines.Quantity = objDoc.Lines.Quantity
            objDoc.Lines.WarehouseCode = objDoc.Lines.WarehouseCode



            objDoc.Lines.BatchNumbers.BaseLineNumber = 0

            objDoc.Lines.BatchNumbers.InternalSerialNumber = "REL1234"
            objDoc.Lines.BatchNumbers.ManufacturerSerialNumber = "REL1234"
            objDoc.Lines.BatchNumbers.Quantity = 10
            objDoc.Lines.BatchNumbers.BatchNumber = "REL1234"
            objDoc.Lines.BatchNumbers.Add()

            objDoc.SaveXML("C:\GRPO.xml")
            If objDoc.Update <> 0 Then
                MsgBox(objAddOn.objCompany.GetLastErrorDescription)
            Else
                MsgBox(objAddOn.objCompany.GetLastErrorDescription)
                MsgBox("Updated")
            End If

        End If



    End Sub
End Class