Imports System.IO
Imports SAPbouiCOM.Framework
Public Class clsAddOn
    Public WithEvents objApplication As SAPbouiCOM.Application
    Public objCompany As SAPbobsCOM.Company
    Dim oProgBarx As SAPbouiCOM.ProgressBar
    Public objGenFunc As Mukesh.SBOLib.GeneralFunctions
    Public objUIXml As Mukesh.SBOLib.UIXML
    Public ZB_row As Integer = 0
    Public SOMenuID As String = "0"
   
    Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD
    Dim oUserTablesMD As SAPbobsCOM.UserTablesMD
    Dim ret As Long
    Dim str As String
    Dim objForm As SAPbouiCOM.Form
    Dim MenuCount As Integer = 0
    Public objQC As clsQC
    Public objRejDetails As clsRejDetails
    Public objInventoryTrans As clsInventoryTransfer
    Public HANA As Boolean = True
    'Public HANA As Boolean = False

    Public HWKEY() As String = New String() {"L1653539483", "H1397589148"}
    Private Sub CheckLicense()

    End Sub
    Function isValidLicense() As Boolean
        Try
            Try
                If objApplication.Forms.ActiveForm.TypeCount > 0 Then
                    For i As Integer = 0 To objApplication.Forms.ActiveForm.TypeCount - 1
                        objApplication.Forms.ActiveForm.Close()
                    Next
                End If
            Catch ex As Exception
            End Try
            objApplication.Menus.Item("257").Activate()
            Dim CrrHWKEY As String = objApplication.Forms.ActiveForm.Items.Item("79").Specific.Value.ToString.Trim
            objApplication.Forms.ActiveForm.Close()

            For i As Integer = 0 To HWKEY.Length - 1
                If HWKEY(i).Trim = CrrHWKEY.Trim Then
                    Return True
                End If
            Next
            MsgBox("Add-on installation failed due to license mismatch", MsgBoxStyle.OkOnly, "License Management")
            Return False
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return True
    End Function

    Public Sub Intialize(ByVal args() As String)
        Try
            Dim oapplication As Application
            If (args.Length < 1) Then oapplication = New Application Else oapplication = New Application(args(0))
            objApplication = Application.SBO_Application
            If isValidLicense() Then
                objApplication.StatusBar.SetText("Establishing Company Connection Please Wait...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                objCompany = Application.SBO_Application.Company.GetDICompany()
                Try
                    createObjects()
                    createTables()
                    createUDOs()
                    loadMenu()
                Catch ex As Exception
                    objAddOn.objApplication.MessageBox(ex.Message)
                    End
                End Try
                objApplication.StatusBar.SetText("Addon Connected Successfully..!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                oapplication.Run()
            Else
                objApplication.StatusBar.SetText("Addon Disconnected due to license mismatch..!!!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            'System.Windows.Forms.Application.Run()
        Catch ex As Exception
            objAddOn.objApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub

    Public Sub Intialize()
        Dim objSBOConnector As New Mukesh.SBOLib.SBOConnector
        objApplication = objSBOConnector.GetApplication(System.Environment.GetCommandLineArgs.GetValue(1))
        objCompany = objSBOConnector.GetCompany(objApplication)
        Try

            createObjects()
            createTables()
            createUDOs()
            'addJobCardReporttype()
            loadMenu()
        Catch ex As Exception
            objAddOn.objApplication.MessageBox(ex.Message)
            End
        End Try
        If isValidLicense() Then
            objApplication.SetStatusBarMessage("Addon connected  successfully!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
        Else
            objApplication.SetStatusBarMessage("Failed To Connect, Please Check The License Configuration", SAPbouiCOM.BoMessageTime.bmt_Short, False)
            objCompany.Disconnect()
            objApplication = Nothing
            objCompany = Nothing
            End
        End If
    End Sub

    Private Sub createUDOs()
        Dim objUDFEngine As New Mukesh.SBOLib.UDFEngine(objCompany)
        Dim ct1(1) As String
        ct1(0) = ""
        objUDFEngine.createUDO("QCWHSE", "QCWHSE", "QCWHSE", ct1, SAPbobsCOM.BoUDOObjType.boud_MasterData, False, False)

        ct1(0) = ""
        objUDFEngine.createUDO("MIPLREJ", "MIPLREJ", "MIPLREJ", ct1, SAPbobsCOM.BoUDOObjType.boud_MasterData, False, False)

        ct1(0) = "MIPLQC1"

        objUDFEngine.createUDO("MIPLQC", "MIPLQC", "MIPLQC", ct1, SAPbobsCOM.BoUDOObjType.boud_Document, False, True)
        ct1(0) = "MIREJDET1"
        'MIREJDET
        objUDFEngine.createUDO("MIREJDET", "MIREJDET", "MIREJDET1", ct1, SAPbobsCOM.BoUDOObjType.boud_Document, False, True)

    End Sub
    Private Sub createObjects()
        objGenFunc = New Mukesh.SBOLib.GeneralFunctions(objCompany)
        objUIXml = New Mukesh.SBOLib.UIXML(objApplication)
        objQC = New clsQC
        objRejDetails = New clsRejDetails
        objInventoryTrans = New clsInventoryTransfer
    End Sub
    Private Sub objApplication_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles objApplication.ItemEvent
        Try
            Select Case pVal.FormTypeEx
                Case clsQC.Formtype
                    objQC.ItemEvent(FormUID, pVal, BubbleEvent)
                Case clsRejDetails.Formtype
                    objRejDetails.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "940" 'inventory transfer
                    objInventoryTrans.ItemEvent(FormUID, pVal, BubbleEvent)
                    'Case "149" 'sales Quotation
                    '    objSQ.ItemEvent(FormUID, pVal, BubbleEvent)
            End Select
        Catch ex As Exception
            objAddOn.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
        End Try
    End Sub
    Private Sub objApplication_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles objApplication.FormDataEvent
        Try
            If BusinessObjectInfo.BeforeAction = True Then
                Select Case BusinessObjectInfo.FormTypeEx
                    Case clsQC.Formtype
                        objQC.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                    Case clsRejDetails.Formtype
                        objRejDetails.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                End Select
            End If

        Catch ex As Exception
            'objAddOn.objApplication.StatusBar.SetText(ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        End Try
    End Sub
    Public Shared Function ErrorHandler(ByVal p_ex As Exception, ByVal objApplication As SAPbouiCOM.Application)
        Dim sMsg As String = Nothing
        If p_ex.Message = "Form - already exists [66000-11]" Then
            Return True
            Exit Function  'ignore error
        End If
        Return False
    End Function
    Private Sub objApplication_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles objApplication.MenuEvent
        If pVal.BeforeAction Then
            Select Case pVal.MenuUID

            End Select
        Else
            Try
                Select Case pVal.MenuUID
                    Case clsQC.Formtype
                        objQC.LoadScreen()
                    Case "1282" ', "1290", "1288", "1289", "1291"
                        If objApplication.Forms.ActiveForm.UniqueID.Contains(clsQC.Formtype) Then
                            objQC.LoadSeries(objApplication.Forms.ActiveForm.UniqueID)
                        End If
                    Case "ditem"
                        objRejDetails.MenuEvent(pVal, BubbleEvent)
                    Case "1290", "1288", "1289", "1291"
                        If objApplication.Forms.ActiveForm.UniqueID.Contains(clsQC.Formtype) Then
                            objQC.TypeSelection(objApplication.Forms.ActiveForm.UniqueID)

                        End If
                    Case "1281"
                        objQC.MenuEvent(pVal, BubbleEvent)
                    Case "1293"
                        objQC.DeleteRow()
                End Select
            Catch ex As Exception
                ' MsgBox(ex.ToString)
            End Try
        End If
        'If pVal.BeforeAction Then
        '    Select Case pVal.MenuUID

        '    End Select
        'Else
        '    Try
        '        Select Case pVal.MenuUID
        '            Case clsQC.Formtype
        '                objQC.LoadScreen()
        '            Case "1282" ', "1290", "1288", "1289", "1291"
        '                If objApplication.Forms.ActiveForm.UniqueID.Contains(clsQC.Formtype) Then
        '                    objQC.LoadSeries(objApplication.Forms.ActiveForm.UniqueID)
        '                End If
        '            Case "ditem"
        '                objRejDetails.MenuEvent(pVal, BubbleEvent)
        '            Case "1290", "1288", "1289", "1291"
        '                If objApplication.Forms.ActiveForm.UniqueID.Contains(clsQC.Formtype) Then
        '                    objQC.TypeSelection(objApplication.Forms.ActiveForm.UniqueID)
        '                End If
        '        End Select
        '    Catch ex As Exception
        '        ' MsgBox(ex.ToString)
        '    End Try
        'End If
    End Sub
    Private Sub loadMenu()
        If objApplication.Menus.Item("43520").SubMenus.Exists("QC") Then Return
        MenuCount = objApplication.Menus.Item("43520").SubMenus.Count

        CreateMenu(Windows.Forms.Application.StartupPath + "\Quality.jpg", MenuCount + 1, "QC Management", SAPbouiCOM.BoMenuType.mt_POPUP, "QC", objApplication.Menus.Item("43520"))
        CreateMenu("", 1, "Quality Check", SAPbouiCOM.BoMenuType.mt_STRING, clsQC.Formtype, objApplication.Menus.Item("QC"))

    End Sub
    Private Function CreateMenu(ByVal ImagePath As String, ByVal Position As Int32, ByVal DisplayName As String, ByVal MenuType As SAPbouiCOM.BoMenuType, ByVal UniqueID As String, ByVal ParentMenu As SAPbouiCOM.MenuItem) As SAPbouiCOM.MenuItem
        Try
            Dim oMenuPackage As SAPbouiCOM.MenuCreationParams
            oMenuPackage = objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
            oMenuPackage.Image = ImagePath
            oMenuPackage.Position = Position
            oMenuPackage.Type = MenuType
            oMenuPackage.UniqueID = UniqueID
            oMenuPackage.String = DisplayName
            ParentMenu.SubMenus.AddEx(oMenuPackage)
        Catch ex As Exception
            objApplication.StatusBar.SetText("Menu Already Exists", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
        End Try
        Return ParentMenu.SubMenus.Item(UniqueID)
    End Function
    Private Sub createTables()
        Dim objUDFEngine As New Mukesh.SBOLib.UDFEngine(objCompany)
        objAddOn.objApplication.SetStatusBarMessage("Creating Tables Please Wait...", SAPbouiCOM.BoMessageTime.bmt_Long, False)
        ' WriteSMSLog("0")
        objUDFEngine.CreateTable("MIQCREJ", "Rejection Details", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objUDFEngine.AddAlphaField("@MIQCREJ", "IssType", "Issue Type", 30)
        objUDFEngine.AddDefaultFormUDO("MIQCREJ", "Rejection Details", SAPbobsCOM.BoUDOObjType.boud_MasterData, "MIQCREJ", {""}, {"Code", "Name"}, True, {"Code", "Name", "U_IssType"}, True, False)

        objUDFEngine.CreateTable("MIPLQC", "QC Header", SAPbobsCOM.BoUTBTableType.bott_Document)
        objUDFEngine.AddAlphaField("@MIPLQC", "Location", "Location", 10)
        objUDFEngine.AddAlphaField("@MIPLQC", "Type", "Type", 10, "-,G,P,T,R", "-,GRN,Production,Transfer,Receipt", "-") ' goods receipt is included for Kanagavalli
        objUDFEngine.AddAlphaField("@MIPLQC", "GRNNum", "Doc Num", 10)
        objUDFEngine.AddAlphaField("@MIPLQC", "GRNEntry", "Doc Entry", 10)
        objUDFEngine.AddAlphaField("@MIPLQC", "ProdNum", "Production Number", 10)
        objUDFEngine.AddAlphaField("@MIPLQC", "ProdEntry", "Production Entry", 10)
        objUDFEngine.AddAlphaField("@MIPLQC", "TransNum", "Transfer Number", 10)
        objUDFEngine.AddAlphaField("@MIPLQC", "TransEntry", "Transfer Entry", 10)
        objUDFEngine.AddDateField("@MIPLQC", "DocDate", "DocDate", SAPbobsCOM.BoFldSubTypes.st_None)
        objUDFEngine.AddAlphaField("@MIPLQC", "InWhse", "InWhse", 30)
        objUDFEngine.AddAlphaField("@MIPLQC", "ApprCode", "ApprCode", 30)
        objUDFEngine.AddAlphaField("@MIPLQC", "ApprName", "ApprName", 100)
        objUDFEngine.AddAlphaField("@MIPLQC", "StkPost", "Stock Posting", 10, "Y,N", "Yes,No", "N")
        objUDFEngine.AddAlphaField("@MIPLQC", "AccStk", "Accepted DocNum", 15)
        objUDFEngine.AddAlphaField("@MIPLQC", "RejStk", "Rejected DocNum", 15)
        objUDFEngine.AddAlphaField("@MIPLQC", "RewStk", "Tework DocNum", 15)
        objUDFEngine.AddAlphaField("@MIPLQC", "Vendor", "Vendor", 100)
        objUDFEngine.AddNumericField("@MIPLQC", "BPLId", "BPLId", 10)
        objUDFEngine.AddAlphaField("@MIPLQC", "BPLName", "BPLName", 100)
        objUDFEngine.AddAlphaField("@MIPLQC", "GREntry", "GR Entry", 10)
        objUDFEngine.AddAlphaField("@MIPLQC", "GRNum", "GR Number", 10)

        objUDFEngine.CreateTable("MIPLQC1", "QC Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objUDFEngine.AddAlphaField("@MIPLQC1", "BaseLinNum", "BaseLinNum", 10)
        objUDFEngine.AddAlphaField("@MIPLQC1", "ItemCode", "Item Code", 30)
        objUDFEngine.AddAlphaField("@MIPLQC1", "ItemName", "Item Name", 100)
        objUDFEngine.AddFloatField("@MIPLQC1", "TotQty", "Total Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objUDFEngine.AddFloatField("@MIPLQC1", "InspQty", "Inspected Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objUDFEngine.AddFloatField("@MIPLQC1", "PendQty", "Pending Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objUDFEngine.AddFloatField("@MIPLQC1", "AccQty", "Accepted Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objUDFEngine.AddFloatField("@MIPLQC1", "RejQty", "Rejected Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objUDFEngine.AddFloatField("@MIPLQC1", "RewQty", "Rework Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objUDFEngine.AddFloatField("@MIPLQC1", "QtyInsp", "Qty Inspected", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objUDFEngine.AddAlphaField("@MIPLQC1", "Remarks", "Remarks", 100)
        objUDFEngine.AddAlphaField("@MIPLQC1", "AccWhse", "Accepted Whse", 30)
        objUDFEngine.AddAlphaField("@MIPLQC1", "RejDet", "Rejection Details", 30)
        objUDFEngine.AddAlphaField("@MIPLQC1", "RejWhse", "Reject Whse", 30)
        objUDFEngine.AddFloatField("@MIPLQC1", "RejPer", "Rejection %", SAPbobsCOM.BoFldSubTypes.st_Percentage)
        objUDFEngine.AddAlphaField("@MIPLQC1", "RewWhse", "Rework whse", 30)
        objUDFEngine.AddFloatField("@MIPLQC1", "SmplQty", "Sample Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objUDFEngine.AddDateField("@MIPLQC1", "InspDate", "InspDate", SAPbobsCOM.BoFldSubTypes.st_None)
        objUDFEngine.AddAlphaField("@MIPLQC1", "EmpCode", "EmpCode", 30)
        objUDFEngine.AddAlphaField("@MIPLQC1", "EmpName", "EmpName", 100)
        objUDFEngine.AddAlphaField("@MIPLQC1", "ItGrpNm", "ItGrpNm", 100)
        objUDFEngine.AddAlphaField("@MIPLQC1", "InvUom", "InvUoM", 30)
        objUDFEngine.AddAlphaField("@MIPLQC1", "RejLoc", "Rejection Location", 40, "-,IN-PRO-REJ,IN-MAT-REJ,OUT-PRO-REJ,OUT-MAT-REJ", "-,In-house Process Rejection,In-house Material Rejection,Out-Source Process Rejection,Out-Source Material Rejection", "-")


        objUDFEngine.AddFloatField("@MIPLQC1", "PPM", "PPM", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objUDFEngine.AddFloatField("@MIPLQC1", "ItemCost", "Item Cost", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUDFEngine.AddFloatField("@MIPLQC1", "RejCost", "Rejected Item Cost", SAPbobsCOM.BoFldSubTypes.st_Sum)

        'U_ItGrpNm


        objUDFEngine.CreateTable("QCWHSE", "QC Warehouse Setting", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objUDFEngine.AddAlphaField("@QCWHSE", "Type", "Type", 10, "-,G,P,T", "-,GRN,Production,Transfer", "-")
        objUDFEngine.AddAlphaField("@QCWHSE", "InWhse", "InWhse", 20)
        objUDFEngine.AddAlphaField("@QCWHSE", "AccWhse", "AccWhse", 20)
        objUDFEngine.AddAlphaField("@QCWHSE", "RejWhse", "RejWhse", 20)
        objUDFEngine.AddAlphaField("@QCWHSE", "RewWhse", "RewWhse", 20)

        objUDFEngine.CreateTable("MIPLREJ", "Rejection Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objUDFEngine.AddAlphaField("@MIPLREJ", "IssType", "Issue Type", 30, "PP,CP,EL", "ProductProcess,Coating/Plating,Electronics", "PP")
        objUDFEngine.AddAlphaField("@MIPLREJ", "Reason", "Reason", 150)

        objUDFEngine.CreateTable("MIREJDET", "Rejection Details Header", SAPbobsCOM.BoUTBTableType.bott_Document)
        objUDFEngine.AddAlphaField("@MIREJDET", "DocNum", "DocNum", 30)
        objUDFEngine.AddDateField("@MIREJDET", "DocDate", "DocDate", SAPbobsCOM.BoFldSubTypes.st_None)
        objUDFEngine.AddAlphaField("@MIREJDET", "ItemCode", "ItemCode", 30)
        objUDFEngine.AddAlphaField("@MIREJDET", "ItemName", "ItemName", 30)
        objUDFEngine.AddFloatField("@MIREJDET", "RejQty", "RejQty", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objUDFEngine.CreateTable("MIREJDET1", "Rejection Details Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objUDFEngine.AddAlphaField("@MIREJDET1", "IssType", "Issue Type", 30, "PP,CP,EL", "ProductProcess,Coating/Plating,Electronics", "PP")
        objUDFEngine.AddFloatField("@MIREJDET1", "Qty", "Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objUDFEngine.AddAlphaField("@MIREJDET1", "Reason", "Reason", 150)
        objUDFEngine.AddAlphaField("@MIREJDET1", "Remarks", "Remarks", 100)

        objUDFEngine.AddAlphaField("OWTR", "QCEntry", "QCEntry", 15)
        objUDFEngine.AddAlphaField("OWTR", "GRNEntry", "GRNEntry", 15)
        objUDFEngine.AddAlphaField("OWTR", "ProdEntry", "ProdEntry", 15)
        objUDFEngine.AddAlphaField("OWTR", "GREntry", "GREntry", 15)
        objUDFEngine.AddAlphaField("OITM", "InspReq", "Inspection Required", 5, "Y,N", "Yes,No", "N")
        objUDFEngine.AddAlphaField("PDN1", "InspReq", "Inspection Required", 5, "Y,N", "Yes,No", "N")

        '*******************  Table ******************* START********************************* END
    End Sub
    Private Sub objApplication_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles objApplication.RightClickEvent
        Dim FormType As String = "MIPLQC"
        If eventInfo.BeforeAction Then
            Select Case eventInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK '"MIPLQC1"
                    Select Case FormType
                        Case "MIPLQC"
                            objQC.RightClickEvent(eventInfo, BubbleEvent)
                    End Select
            End Select
        Else

            'If eventInfo.FormUID.Contains("MIREJDET") And (eventInfo.ItemUID = "13") And eventInfo.Row > 0 Then
            '    Dim oMenuItem As SAPbouiCOM.MenuItem
            '    Dim oMenus As SAPbouiCOM.Menus
            '    Try
            '        If objAddOn.objApplication.Menus.Exists("ditem") Then
            '            objAddOn.objApplication.Menus.RemoveEx("ditem")
            '        End If
            '    Catch ex As Exception
            '    End Try
            '    Try
            '        oMenuItem = objAddOn.objApplication.Menus.Item("1280").SubMenus.Item("ditem")
            '        ZB_row = eventInfo.Row
            '    Catch ex As Exception
            '        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
            '        oCreationPackage = objAddOn.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
            '        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            '        oCreationPackage.UniqueID = "ditem"
            '        oCreationPackage.String = "Delete Row"
            '        oCreationPackage.Enabled = True
            '        oMenuItem = objAddOn.objApplication.Menus.Item("1280") 'Data'
            '        oMenus = oMenuItem.SubMenus
            '        oMenus.AddEx(oCreationPackage)
            '        ZB_row = eventInfo.Row
            '    End Try
            '    If eventInfo.ItemUID <> "13" Then
            '        '   Dim oMenuItem As SAPbouiCOM.MenuItem
            '        '  Dim oMenus As SAPbouiCOM.Menus
            '        Try
            '            objAddOn.objApplication.Menus.RemoveEx("ditem")
            '        Catch ex As Exception
            '            ' MessageBox.Show(ex.Message)
            '        End Try
            '    End If
            'End If
        End If
    End Sub

    Private Sub objApplication_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles objApplication.AppEvent
        If EventType = SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged Or EventType = SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition Or EventType = SAPbouiCOM.BoAppEventTypes.aet_ShutDown Then
            Try
                ' objUIXml.LoadMenuXML("RemoveMenu.xml", Mukesh.SBOLib.UIXML.enuResourceType.Embeded)
                If objCompany.Connected Then objCompany.Disconnect()
                objCompany = Nothing
                objApplication = Nothing
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objCompany)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objApplication)
                GC.Collect()
            Catch ex As Exception
            End Try
            End
        End If
    End Sub

    Private Sub applyFilter()
        Dim oFilters As SAPbouiCOM.EventFilters
        Dim oFilter As SAPbouiCOM.EventFilter
        oFilters = New SAPbouiCOM.EventFilters
        'Item Master Data 
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_GOT_FOCUS)

        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK)

        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE)

        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK)


    End Sub
    Public Sub WriteSMSLog(ByVal Str As String)
        Dim fs As FileStream
        Dim chatlog As String = System.Windows.Forms.Application.StartupPath & "\Log_" & Today.ToString("yyyyMMdd") & ".txt"
        If File.Exists(chatlog) Then
        Else
            fs = New FileStream(chatlog, FileMode.Create, FileAccess.Write)
            fs.Close()
        End If
        ' Dim objReader As New System.IO.StreamReader(chatlog)
        Dim sdate As String
        sdate = Now
        'objReader.Close()
        If System.IO.File.Exists(chatlog) = True Then
            Dim objWriter As New System.IO.StreamWriter(chatlog, True)
            objWriter.WriteLine(sdate & " : " & Str)
            objWriter.Close()
        Else
            Dim objWriter As New System.IO.StreamWriter(chatlog, False)
            ' MsgBox("Failed to send message!")
        End If
    End Sub
    Private Sub addJobCardReporttype()
        'Dim rptTypeService As SAPbobsCOM.ReportTypesService
        'Dim newType As SAPbobsCOM.ReportType
        'Dim newtypeParam As SAPbobsCOM.ReportTypeParams
        'Dim newReportParam As SAPbobsCOM.ReportLayoutParams
        'Dim ReportExists As Boolean = False
        'Try

        '    Dim newtypesParam As SAPbobsCOM.ReportTypesParams
        '    rptTypeService = objAddOn.objCompany.GetCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.ReportTypesService)
        '    newtypesParam = rptTypeService.GetReportTypeList

        '    Dim i As Integer
        '    For i = 0 To newtypesParam.Count - 1
        '        If newtypesParam.Item(i).TypeName = clsJobCard.FormType And newtypesParam.Item(i).MenuID = clsJobCard.FormType Then
        '            ReportExists = True
        '            Exit For
        '        End If
        '    Next i

        '    If Not ReportExists Then
        '        rptTypeService = objCompany.GetCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.ReportTypesService)
        '        newType = rptTypeService.GetDataInterface(SAPbobsCOM.ReportTypesServiceDataInterfaces.rtsReportType)


        '        newType.TypeName = clsJobCard.FormType
        '        newType.AddonName = "JC2Addon"
        '        newType.AddonFormType = clsJobCard.FormType
        '        newType.MenuID = clsJobCard.FormType
        '        newtypeParam = rptTypeService.AddReportType(newType)

        '        Dim rptService As SAPbobsCOM.ReportLayoutsService
        '        Dim newReport As SAPbobsCOM.ReportLayout
        '        rptService = objCompany.GetCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.ReportLayoutsService)
        '        newReport = rptService.GetDataInterface(SAPbobsCOM.ReportLayoutsServiceDataInterfaces.rlsdiReportLayout)
        '        newReport.Author = objCompany.UserName
        '        newReport.Category = SAPbobsCOM.ReportLayoutCategoryEnum.rlcCrystal
        '        newReport.Name = clsJobCard.FormType
        '        newReport.TypeCode = newtypeParam.TypeCode

        '        newReportParam = rptService.AddReportLayout(newReport)

        '        newType = rptTypeService.GetReportType(newtypeParam)
        '        newType.DefaultReportLayout = newReportParam.LayoutCode
        '        rptTypeService.UpdateReportType(newType)

        '        Dim oBlobParams As SAPbobsCOM.BlobParams
        '        oBlobParams = objCompany.GetCompanyService.GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiBlobParams)
        '        oBlobParams.Table = "RDOC"
        '        oBlobParams.Field = "Template"
        '        Dim oKeySegment As SAPbobsCOM.BlobTableKeySegment
        '        oKeySegment = oBlobParams.BlobTableKeySegments.Add
        '        oKeySegment.Name = "DocCode"
        '        oKeySegment.Value = newReportParam.LayoutCode

        '        Dim oFile As FileStream
        '        oFile = New FileStream(Application.StartupPath + "\JobCard.rpt", FileMode.Open)
        '        Dim fileSize As Integer
        '        fileSize = oFile.Length
        '        Dim buf(fileSize) As Byte
        '        oFile.Read(buf, 0, fileSize)
        '        oFile.Dispose()

        '        Dim oBlob As SAPbobsCOM.Blob
        '        oBlob = objCompany.GetCompanyService.GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiBlob)
        '        oBlob.Content = Convert.ToBase64String(buf, 0, fileSize)
        '        objCompany.GetCompanyService.SetBlob(oBlobParams, oBlob)
        '    End If
        'Catch ex As Exception
        '    objApplication.MessageBox(ex.ToString)
        'End Try

    End Sub

    Private Sub objApplication_LayoutKeyEvent(ByRef eventInfo As SAPbouiCOM.LayoutKeyInfo, ByRef BubbleEvent As Boolean) Handles objApplication.LayoutKeyEvent

        ''BubbleEvent = True
        'If eventInfo.BeforeAction = True Then
        '    If eventInfo.FormUID.Contains(clsJobCard.FormType) Then
        '        objJobCard.LayoutKeyEvent(eventInfo, BubbleEvent)
        '    End If
        'End If
    End Sub

    Private Sub objApplication_ProgressBarEvent(ByRef pVal As SAPbouiCOM.ProgressBarEvent, ByRef BubbleEvent As Boolean) Handles objApplication.ProgressBarEvent

    End Sub
End Class


