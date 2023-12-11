Imports SAPbouiCOM
Namespace SubContractingPO

    Public Class clsMenuEvent
        Dim objform As SAPbouiCOM.Form
        Dim objglobalmethods As New clsGlobalMethods

        Public Sub MenuEvent_For_StandardMenu(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Try
                Select Case objaddon.objapplication.Forms.ActiveForm.TypeEx
                    Case "SUBCTPO"
                        SubContractingPO_MenuEvent(pVal, BubbleEvent)
                    Case "SUBBOM"
                        SubContractingBOM_MenuEvent(pVal, BubbleEvent)
                        'Case "SUBGEN"
                        '    GeneralSettings_MenuEvent(pVal, BubbleEvent)
                    Case "65211"
                        ProductionOrder_MenuEvent(pVal, BubbleEvent)
                    Case "940", "721", "720", "-940", "-721", "-720"
                        Default_Sample_MenuEvent(pVal, BubbleEvent)
                End Select
            Catch ex As Exception

            End Try
        End Sub

        Private Sub Default_Sample_MenuEvent(ByVal pval As SAPbouiCOM.MenuEvent, ByVal BubbleEvent As Boolean)
            Try
                objform = objaddon.objapplication.Forms.ActiveForm
                If pval.BeforeAction = True Then
                Else
                    Dim oUDFForm As SAPbouiCOM.Form
                    Try
                        oUDFForm = objaddon.objapplication.Forms.Item(objform.UDFFormUID)
                    Catch ex As Exception
                        oUDFForm = objform
                    End Try

                    Select Case pval.MenuUID
                        Case "1281"
                            oUDFForm.Items.Item("U_SubConNo").Enabled = True
                        Case "1287"
                            If oUDFForm.Items.Item("U_SubConNo").Enabled = False Then
                                oUDFForm.Items.Item("U_SubConNo").Enabled = True
                            End If
                            oUDFForm.Items.Item("U_SubConNo").Specific.String = ""

                        Case Else
                            oUDFForm.Items.Item("U_SubConNo").Enabled = False
                    End Select
                End If
            Catch ex As Exception
                objaddon.objapplication.SetStatusBarMessage("Error in Standart Menu Event" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try
        End Sub

#Region "SubContractingPO"
      
        Private Sub SubContractingPO_MenuEvent(ByRef pval As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Dim Matrix0, Matrix2, Matrix4, Matrix3, Matrix1 As SAPbouiCOM.Matrix
            Dim FolderInput, FolderOutput, FolderScrap, FolderRelDoc, FolderCosting As SAPbouiCOM.Folder
            Dim FolderID As String = ""
            Try
                objform = objaddon.objapplication.Forms.ActiveForm
                Matrix0 = objform.Items.Item("MtxinputN").Specific
                Matrix2 = objform.Items.Item("mtxreldoc").Specific
                Matrix4 = objform.Items.Item("MtxCosting").Specific
                Matrix3 = objform.Items.Item("mtxoutput").Specific
                Matrix1 = objform.Items.Item("mtxscrap").Specific
                FolderInput = objform.Items.Item("flrinput").Specific
                FolderOutput = objform.Items.Item("flroutput").Specific
                FolderScrap = objform.Items.Item("flrscrap").Specific
                FolderRelDoc = objform.Items.Item("flrreldoc").Specific
                FolderCosting = objform.Items.Item("flrcosting").Specific
                If pval.BeforeAction = True Then
                    Select Case pval.MenuUID
                        Case "1283", "1284" 'Remove & Cancel
                            objaddon.objapplication.SetStatusBarMessage("Remove or Cancel is not allowed ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        Case "1293"  'Delete Row
                        Case "6913"
                            If objform.TypeEx = "SUBCTPO" Then
                                BubbleEvent = False
                            End If

                    End Select
                Else
                    If FolderInput.Selected = True Then
                        FolderID = "MtxinputN"
                    ElseIf FolderOutput.Selected = True Then
                        FolderID = "mtxoutput"
                    ElseIf FolderScrap.Selected = True Then
                        FolderID = "mtxscrap"
                    ElseIf FolderRelDoc.Selected = True Then
                        FolderID = "mtxreldoc"
                    ElseIf FolderCosting.Selected = True Then
                        FolderID = "MtxCosting"
                    End If
                    Dim DBSource As SAPbouiCOM.DBDataSource
                    DBSource = objform.DataSources.DBDataSources.Item("@MIPL_OPOR")
                    Select Case pval.MenuUID
                        Case "1281" 'Find Mode
                            objform.Items.Item("txtdocnum").Enabled = True
                            objform.Items.Item("txtentry").Enabled = True
                            objform.Items.Item("txtstat").Enabled = True
                            objform.Items.Item("txtcode").Enabled = True
                            objform.Items.Item("posdate").Enabled = True
                            objform.Items.Item("txtsitem").Enabled = True
                            'objform.Items.Item("txtbitem").Enabled = True
                            objform.Items.Item("docdate").Enabled = True
                            objform.Items.Item("deldate").Enabled = True
                            objform.Items.Item("txtentry").Enabled = True
                            objform.Items.Item("txtGINo").Enabled = True
                            objform.Items.Item("txtGRNo").Enabled = True
                            objform.Items.Item("TxtInvTr").Enabled = True
                            objform.Items.Item("tpodoc").Enabled = True
                            objform.Items.Item("tpoentry").Enabled = True
                            Matrix0.Item.Enabled = False
                            Matrix1.Item.Enabled = False
                            Matrix2.Item.Enabled = False
                            Matrix3.Item.Enabled = False
                            Matrix4.Item.Enabled = False
                            'objform.ActiveItem = "txtdocnum"
                            objform.Items.Item("txtdocnum").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            If objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                objform = objaddon.objapplication.Forms.GetForm("SUBCTPO", 0)
                            End If
                            Exit Sub
                        Case "1282" ' Add Mode
                            objform.Items.Item("btngendoc").Enabled = False
                            objform.Items.Item("btnload").Enabled = False
                            objform.Items.Item("BtnView").Enabled = False
                            objform.Items.Item("BtnInv").Enabled = False
                            objform.Items.Item("BtnInvO").Enabled = False
                            objform.Items.Item("AutoPO").Enabled = False
                            objform.Items.Item("BtnGIssue").Enabled = False
                            objform.Items.Item("btnOutput").Enabled = False
                            objform.Items.Item("BtnScrap").Enabled = False
                            objform.Items.Item("InvScrap").Enabled = False
                            objform.Items.Item("posdate").Specific.string = Now.Date.ToString("yyyyMMdd")
                            objform.Items.Item("docdate").Specific.string = Now.Date.ToString("yyyyMMdd")
                            objform.Items.Item("txtremark").Specific.String = "Created By " & objaddon.objcompany.UserName & " on " & Now.ToString

                            objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "Code", "#")
                            objaddon.objglobalmethods.Matrix_Addrow(Matrix2, "Code", "#")
                            objaddon.objglobalmethods.Matrix_Addrow(Matrix4, "Code", "#")
                            objaddon.objglobalmethods.Matrix_Addrow(Matrix3, "Code", "#")
                            objaddon.objglobalmethods.Matrix_Addrow(Matrix1, "Code", "#")
                            FolderScrap.Item.Click(BoCellClickType.ct_Regular)
                            FolderInput.Item.Click(BoCellClickType.ct_Regular)
                            objaddon.objglobalmethods.LoadSeries(objform, DBSource)
                            'objform.Items.Item("txtdocnum").Specific.string = objaddon.objglobalmethods.GetNextDocNum_Value("@MIPL_OPOR")
                            'objform.Items.Item("txtentry").Specific.string = objaddon.objglobalmethods.GetNextDocEntry_Value("@MIPL_OPOR")
                            For i As Integer = 1 To Matrix1.VisualRowCount
                                If Matrix1.Columns.Item("Code").Cells.Item(i).Specific.String = "" And Matrix1.Columns.Item("Colstat").Cells.Item(i).Specific.String = "O" Then
                                    Matrix1.Columns.Item("Code").Editable = True
                                    Matrix1.Columns.Item("Quant").Editable = True
                                    Matrix1.Columns.Item("Whse").Editable = True
                                End If
                            Next
                        Case "1288", "1289", "1290", "1291"
                        Case "1287"  'Duplicate
                            Matrix3.ClearRowData(Matrix3.VisualRowCount)
                            Matrix2.ClearRowData(Matrix2.VisualRowCount)
                            Matrix4.ClearRowData(Matrix4.VisualRowCount)
                            Matrix1.ClearRowData(Matrix1.VisualRowCount)
                            Matrix3.Clear()
                            Matrix4.Clear()
                            Matrix2.Clear()
                            Matrix1.Clear()
                            objform.Items.Item("txtGINo").Specific.String = ""
                            objform.Items.Item("txtGRNo").Specific.String = ""
                            objform.Items.Item("TxtInvTr").Specific.String = ""
                            objform.Items.Item("txtstat").Specific.String = "O"
                            objform.Items.Item("txtentry").Specific.String = ""
                            objform.Items.Item("txtopqty").Specific.String = ""
                            objform.Items.Item("txtvobstk").Specific.String = ""
                            objform.Items.Item("txtvbal").Specific.String = ""
                            objform.Items.Item("deldate").Specific.String = ""
                            objaddon.objglobalmethods.LoadSeries(objform, DBSource)
                            objform.Items.Item("posdate").Specific.string = Now.Date.ToString("yyyyMMdd")
                            objform.Items.Item("docdate").Specific.string = Now.Date.ToString("yyyyMMdd")
                            objform.Items.Item("txtremark").Specific.String = "Created By " & objaddon.objcompany.UserName & " on " & Now.ToString
                            objform.Items.Item("btngendoc").Enabled = False
                            objform.Items.Item("btnload").Enabled = False
                            objform.Items.Item("BtnView").Enabled = False
                            objform.Items.Item("BtnInv").Enabled = False
                            objform.Items.Item("BtnGIssue").Enabled = False
                            FolderOutput.Item.Click(BoCellClickType.ct_Regular)
                            objform.Items.Item("btnOutput").Enabled = False
                            objform.Items.Item("BtnScrap").Enabled = False
                            objform.Items.Item("InvScrap").Enabled = False
                            objform.Items.Item("AutoPO").Enabled = False
                            objform.Items.Item("BtnInvO").Enabled = False
                            objform.Items.Item("txtponum").Specific.String = ""
                            objform.Items.Item("tpoentry").Specific.String = ""
                            objform.Items.Item("SubPoNum").Specific.String = ""
                            objform.Items.Item("tpodoc").Specific.String = ""
                            objform.Items.Item("POLine").Specific.String = ""
                            FolderInput.Item.Click(BoCellClickType.ct_Regular)
                        Case "1293"
                            Select Case FolderID
                                Case "MtxinputN"
                                    DeleteRow(Matrix0, "@MIPL_POR1")
                                Case "mtxoutput"
                                    DeleteRow(Matrix3, "@MIPL_POR2")
                                Case "mtxscrap"
                                    DeleteRow(Matrix1, "@MIPL_POR3")
                                Case "mtxreldoc"
                                    DeleteRow(Matrix2, "@MIPL_POR4")
                                Case "MtxCosting"
                                    DeleteRow(Matrix4, "@MIPL_POR5")
                            End Select
                        Case "1292"
                            Select Case FolderID
                                Case "MtxinputN"
                                    objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "Code", "#")
                                Case "mtxoutput"
                                    objaddon.objglobalmethods.Matrix_Addrow(Matrix3, "Code", "#")
                                    Matrix3.CommonSetting.SetCellEditable(Matrix3.VisualRowCount, 2, True)
                                    Matrix3.CommonSetting.SetCellEditable(Matrix3.VisualRowCount, 4, True)
                                    Matrix3.CommonSetting.SetCellEditable(Matrix3.VisualRowCount, 10, True)
                                    objaddon.objglobalmethods.GetSubItemQty(objform, Matrix3)
                                Case "mtxscrap"
                                    objaddon.objglobalmethods.Matrix_Addrow(Matrix1, "Code", "#")
                                Case "mtxreldoc"
                                    objaddon.objglobalmethods.Matrix_Addrow(Matrix2, "Code", "#")
                                Case "MtxCosting"
                                    objaddon.objglobalmethods.Matrix_Addrow(Matrix4, "Code", "#")
                            End Select
                        Case "1304" 'Refresh
                            Dim OnHand As String
                            objform.Mode = BoFormMode.fm_UPDATE_MODE
                            Select Case FolderID
                                Case "MtxinputN"
                                    For i As Integer = 1 To Matrix0.VisualRowCount
                                        If objaddon.HANA Then
                                            OnHand = objaddon.objglobalmethods.getSingleValue("select ""OnHand"" from  oitw where ""ItemCode""='" & Matrix0.Columns.Item("Code").Cells.Item(i).Specific.String & "' and ""WhsCode""='" & Matrix0.Columns.Item("Whse").Cells.Item(i).Specific.String & "'")
                                        Else
                                            OnHand = objaddon.objglobalmethods.getSingleValue("select OnHand from  oitw where ItemCode='" & Matrix0.Columns.Item("Code").Cells.Item(i).Specific.String & "' and WhsCode='" & Matrix0.Columns.Item("Whse").Cells.Item(i).Specific.String & "'")
                                        End If
                                        Matrix0.Columns.Item("Instock").Cells.Item(i).Specific.String = OnHand
                                    Next
                            End Select
                            If objform.Mode = BoFormMode.fm_UPDATE_MODE Then
                                objform.Items.Item("1").Click()
                            End If
                    End Select
                End If
            Catch ex As Exception
                objform.Freeze(False)
                ' objaddon.objapplication.SetStatusBarMessage("Error in Menu Event" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End Try
        End Sub

        Sub DeleteRow(ByVal objMatrix As SAPbouiCOM.Matrix, ByVal TableName As String)
            Try
                Dim DBSource As SAPbouiCOM.DBDataSource
                'objMatrix = objform.Items.Item("20").Specific
                objMatrix.FlushToDataSource()
                DBSource = objform.DataSources.DBDataSources.Item(TableName) '"@MIREJDET1"
                For i As Integer = 1 To objMatrix.VisualRowCount
                    objMatrix.GetLineData(i)
                    DBSource.Offset = i - 1
                    DBSource.SetValue("LineId", DBSource.Offset, i)
                    objMatrix.SetLineData(i)
                    objMatrix.FlushToDataSource()
                Next
                DBSource.RemoveRecord(DBSource.Size - 1)
                objMatrix.LoadFromDataSource()

            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText("DeleteRow  Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
            Finally
            End Try
        End Sub
#End Region

        Private Sub SubContractingBOM_MenuEvent(ByRef pval As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Dim Matrix0 As SAPbouiCOM.Matrix
            Dim odbdsDetails As SAPbouiCOM.DBDataSource

            Try
                objform = objaddon.objapplication.Forms.ActiveForm
                odbdsDetails = objform.DataSources.DBDataSources.Item("@MIPL_BOM1")
                Matrix0 = objform.Items.Item("mtxBOM").Specific
                If pval.BeforeAction = True Then
                    Select Case pval.MenuUID
                        Case "1283", "1284" 'Remove & Cancel
                            'objaddon.objapplication.SetStatusBarMessage("Remove or Cancel is not allowed ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            'BubbleEvent = False
                        Case "1293"
                            If Matrix0.VisualRowCount = 1 Then BubbleEvent = False
                        Case "1292"
                        Case "6913"
                            If objform.TypeEx = "SUBBOM" Then
                                BubbleEvent = False
                            End If
                    End Select
                Else
                    Select Case pval.MenuUID
                        Case "1281" 'Find Mode                           
                            objform.Items.Item("txtentry").Enabled = True
                            objform.Items.Item("txtcode").Enabled = True
                            objform.Items.Item("txtprodno").Enabled = True
                            objform.Items.Item("mtxBOM").Enabled = False
                            objform.Items.Item("txtentry").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        Case "1282" ' Add Mode
                            If objaddon.HANA Then
                                objform.Items.Item("txtentry").Specific.string = objaddon.objglobalmethods.getSingleValue("select Count(*)+1 ""DocEntry"" from ""@MIPL_OBOM""")
                            Else
                                objform.Items.Item("txtentry").Specific.string = objaddon.objglobalmethods.getSingleValue("select Count(*)+1 DocEntry from [@MIPL_OBOM]")
                            End If
                            objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "Code", "#")
                            If ResourceEnable <> "Y" Then
                                Dim cmbtype As SAPbouiCOM.Column = Matrix0.Columns.Item("Type")
                                cmbtype.ValidValues.Remove("290", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            End If
                        Case "1288", "1289", "1290", "1291"
                            'objform.Items.Item("btngendoc").Enabled = True
                            objaddon.objapplication.Menus.Item("1300").Activate()
                        Case "1293"
                            For i As Integer = Matrix0.VisualRowCount To 1 Step -1
                                Matrix0.Columns.Item("#").Cells.Item(i).Specific.String = i
                            Next
                            If objform.Mode = BoFormMode.fm_OK_MODE Then objform.Mode = BoFormMode.fm_UPDATE_MODE
                            objform.Update()
                            objform.Refresh()
                        Case "1292"
                            Try
                                If Matrix0.VisualRowCount > 0 Then
                                    If odbdsDetails.GetValue("U_Itemcode", Matrix0.VisualRowCount - 1) = "" Then Exit Sub
                                    objform.Freeze(True)
                                    odbdsDetails.InsertRecord(odbdsDetails.Size)
                                    odbdsDetails.SetValue("LineId", Matrix0.VisualRowCount, Matrix0.VisualRowCount + 1)
                                    Matrix0.LoadFromDataSource()
                                    objform.Freeze(False)
                                End If
                            Catch ex As Exception

                            End Try
                    End Select
                End If
            Catch ex As Exception
                objform.Freeze(False)
                'objaddon.objapplication.SetStatusBarMessage("Error in Menu Event" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try
        End Sub

        Private Sub GeneralSettings_MenuEvent(ByRef pval As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Try
                objform = objaddon.objapplication.Forms.ActiveForm
                If pval.BeforeAction = True Then
                    Select Case pval.MenuUID
                        Case "1283", "1284" 'Remove & Cancel
                            objaddon.objapplication.SetStatusBarMessage("Remove or Cancel is not allowed ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        Case "1293"

                        Case "1292"

                    End Select
                Else
                    Select Case pval.MenuUID
                        Case "1281" 'Find Mode                           

                        Case "1282" ' Add Mode
                        Case "1288", "1289", "1290", "1291"
                        Case "1293"

                        Case "1292"

                    End Select
                End If
            Catch ex As Exception
                objform.Freeze(False)
                'objaddon.objapplication.SetStatusBarMessage("Error in Menu Event" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try
        End Sub

        Private Sub ProductionOrder_MenuEvent(ByRef pval As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Try
                objform = objaddon.objapplication.Forms.ActiveForm

                If pval.BeforeAction = True Then
                    Select Case pval.MenuUID
                        Case "1283", "1284" 'Remove & Cancel
                            'objaddon.objapplication.SetStatusBarMessage("Remove or Cancel is not allowed ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            'BubbleEvent = False
                    End Select
                Else
                    Select Case pval.MenuUID
                        Case "1281" 'Find Mode 

                        Case "1282" ' Add Mode
                        Case "1288", "1289", "1290", "1291", "1304"
                            If SubPOScreen <> "" Then
                                objaddon.objglobalmethods.SubConButtonEnabling(objform)
                            End If
                        Case "1293"
                        Case "1292"
                    End Select
                End If
            Catch ex As Exception
                objform.Freeze(False)
                objaddon.objapplication.SetStatusBarMessage("Error in Menu Event" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try
        End Sub
     
    End Class
End Namespace