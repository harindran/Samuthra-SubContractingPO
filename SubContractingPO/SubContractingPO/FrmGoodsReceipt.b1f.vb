Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework

Namespace SubContractingPO
    <FormAttribute("721", "SubContractingPO/FrmGoodsReceipt.b1f")>
    Friend Class FrmGoodsReceipt
        Inherits SystemFormBase
        Private WithEvents objSubConform As SAPbouiCOM.Form
        Private WithEvents objform As SAPbouiCOM.Form
        Private WithEvents objUDFform As SAPbouiCOM.Form
        Dim FormCount As Integer = 0
        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.StaticText0 = CType(Me.GetItem("SubConNo").Specific, SAPbouiCOM.StaticText)
            Me.EditText0 = CType(Me.GetItem("U_SubConNo").Specific, SAPbouiCOM.EditText)
            Me.Button0 = CType(Me.GetItem("1").Specific, SAPbouiCOM.Button)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()
            AddHandler DataAddAfter, AddressOf Me.Form_DataAddAfter
            AddHandler DataLoadAfter, AddressOf Me.Form_DataLoadAfter
            AddHandler DeactivateAfter, AddressOf Me.Form_DeactivateAfter
            AddHandler LoadAfter, AddressOf Me.Form_LoadAfter

        End Sub
        Private WithEvents StaticText0 As SAPbouiCOM.StaticText

        Private Sub OnCustomInitialize()
            Try
                'objform = objaddon.objapplication.Forms.GetForm("721", Me.FormCount)
                'objform = objaddon.objapplication.Forms.ActiveForm
                'objSubConform = objaddon.objapplication.Forms.GetForm("SUBCTPO", 0)
            Catch ex As Exception

            End Try
        End Sub
        Private WithEvents EditText0 As SAPbouiCOM.EditText
        Private WithEvents Button0 As SAPbouiCOM.Button
        Private Sub RemoveLastrow(ByVal omatrix As SAPbouiCOM.Matrix, ByVal Columname_check As String)
            Try
                If omatrix.VisualRowCount = 0 Then Exit Sub
                If Columname_check.ToString = "" Then Exit Sub
                If omatrix.Columns.Item(Columname_check).Cells.Item(omatrix.VisualRowCount).Specific.string = "" Then
                    omatrix.DeleteRow(omatrix.VisualRowCount)
                End If
            Catch ex As Exception

            End Try
        End Sub
        Private Sub Button0_ClickBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles Button0.ClickBefore
            Try
                'Dim oUDFForm As SAPbouiCOM.Form
                'oUDFForm = objaddon.objapplication.Forms.Item(objform.UDFFormUID)
                'oUDFForm.Items.Item("U_SubConNo").Enabled = False
                If objform.Items.Item("U_SubConNo").Specific.String = "" Then Exit Sub
                Dim StrQuery As String
                Dim objmatrix As SAPbouiCOM.Matrix
                Dim objRS As SAPbobsCOM.Recordset
                objRS = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objmatrix = objform.Items.Item("13").Specific
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    Dim FolderOutput, FolderScrap As SAPbouiCOM.Folder
                    objSubConform = objaddon.objapplication.Forms.GetForm("SUBCTPO", 1)
                    FolderOutput = objSubConform.Items.Item("flroutput").Specific
                    FolderScrap = objSubConform.Items.Item("flrscrap").Specific
                    objmatrix = objform.Items.Item("13").Specific
                    'RemoveLastrow(objmatrix, "1")
                    Dim Rowcount As Integer = objmatrix.VisualRowCount - 1
                    If FolderOutput.Selected = True Then
                        If objform.Items.Item("U_SubConNo").Specific.String <> "" Then
                            If objaddon.HANA Then
                                StrQuery = "select ""U_Itemcode"" from ""@MIPL_POR2"" where ""DocEntry""=" & objform.Items.Item("U_SubConNo").Specific.String & " and ""U_Status"" ='O'"
                            Else
                                StrQuery = "select U_Itemcode from [@MIPL_POR2] where DocEntry=" & objform.Items.Item("U_SubConNo").Specific.String & " and U_Status ='O'"
                            End If
                            objRS.DoQuery(StrQuery)
                        End If
                        If objRS.RecordCount > 0 And Rowcount > 0 Then
                            If Rowcount <> objRS.RecordCount Then
                                objaddon.objapplication.StatusBar.SetText("Item Count is mismatching with the Output Tab of Sub Contracting.Please check...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False : Exit Sub
                            End If
                        End If
                        For i As Integer = 1 To objmatrix.VisualRowCount
                            If objmatrix.Columns.Item("1").Cells.Item(i).Specific.String <> "" Then
                                If objaddon.HANA Then
                                    StrQuery = "Select distinct 1 from ""@MIPL_POR2"" T2 where T2.""DocEntry""='" & objform.Items.Item("U_SubConNo").Specific.String & "' and T2.""LineId""='" & objmatrix.Columns.Item("U_LineID").Cells.Item(i).Specific.String & "' group by T2.""DocEntry"",T2.""LineId"""
                                    StrQuery += vbCrLf + "having (Select Ifnull(""U_Qty"",0) from ""@MIPL_OPOR"" A join ""@MIPL_POR2"" B on A.""DocEntry""=B.""DocEntry"" where A.""DocEntry""=T2.""DocEntry"" and B.""LineId""=T2.""LineId"")"
                                    StrQuery += vbCrLf + "- ifnull(( Select sum(T1.""Quantity"") from OIGN T0 inner join IGN1 T1 on T0.""DocEntry""=T1.""DocEntry"" where T0.""U_SubConNo""=T2.""DocEntry"""
                                    StrQuery += vbCrLf + "and T1.""U_LineID""=T2.""LineId""),0) <=0"

                                    StrQuery = objaddon.objglobalmethods.getSingleValue(StrQuery)
                                Else
                                    StrQuery = "Select distinct 1 from [@MIPL_POR2] T2 where T2.DocEntry='" & objform.Items.Item("U_SubConNo").Specific.String & "' and T2.LineId='" & objmatrix.Columns.Item("U_LineID").Cells.Item(i).Specific.String & "' group by T2.DocEntry,T2.LineId"
                                    StrQuery += vbCrLf + "having (Select isnull(U_Qty,0) from [@MIPL_OPOR] A join [@MIPL_POR2] B on A.DocEntry=B.DocEntry where A.DocEntry=T2.DocEntry and B.LineId=T2.LineId)"
                                    StrQuery += vbCrLf + "- isnull(( Select sum(T1.Quantity) from OIGN T0 inner join IGN1 T1 on T0.DocEntry=T1.DocEntry where T0.U_SubConNo=T2.DocEntry"
                                    StrQuery += vbCrLf + "and T1.U_LineID=T2.LineId),0) <=0"

                                    StrQuery = objaddon.objglobalmethods.getSingleValue(StrQuery)
                                End If
                                If StrQuery = "1" Then
                                    objaddon.objapplication.StatusBar.SetText("Sub Item Qty Fulfilled.please check...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False : Exit Sub
                                End If
                                If objaddon.HANA Then
                                    objRS.DoQuery("select ""U_Itemcode"",""U_WhsCode"",""U_Qty"",""U_Price"" from ""@MIPL_POR2"" where ""DocEntry""=" & objform.Items.Item("U_SubConNo").Specific.String & " and ""U_Status"" ='O'")
                                Else
                                    objRS.DoQuery("select U_Itemcode,U_WhsCode,U_Qty,U_Price from [@MIPL_POR2] where DocEntry=" & objform.Items.Item("U_SubConNo").Specific.String & " and U_Status ='O'")
                                End If
                                If objRS.RecordCount > 0 Then
                                    'Dim Value As String = objmatrix.Columns.Item("10").Cells.Item(i).Specific.String
                                    'Value = Value.Replace("INR ", "")
                                    If objmatrix.Columns.Item("1").Cells.Item(i).Specific.String <> objRS.Fields.Item("U_Itemcode").Value.ToString Or objmatrix.Columns.Item("15").Cells.Item(i).Specific.String <> objRS.Fields.Item("U_WhsCode").Value.ToString Or CDbl(objmatrix.Columns.Item("9").Cells.Item(i).Specific.String) <> CDbl(objRS.Fields.Item("U_Qty").Value.ToString) Then
                                        objaddon.objapplication.StatusBar.SetText("Line Item is not matching-Output Tab.Please re-generate from Sub Contracting screen.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False : Exit Sub
                                    End If
                                End If
                            End If
                        Next
                    ElseIf FolderScrap.Selected = True Then
                        If objform.Items.Item("U_SubConNo").Specific.String <> "" Then
                            If objaddon.HANA Then
                                StrQuery = "select ""U_Itemcode"" from ""@MIPL_POR3"" where ""DocEntry""=" & objform.Items.Item("U_SubConNo").Specific.String & " and ""U_Status"" ='O'"
                            Else
                                StrQuery = "select U_Itemcode from [@MIPL_POR3] where DocEntry=" & objform.Items.Item("U_SubConNo").Specific.String & " and U_Status ='O'"
                            End If
                            objRS.DoQuery(StrQuery)
                        End If
                        If objRS.RecordCount > 0 And Rowcount > 0 Then
                            If Rowcount <> objRS.RecordCount Then
                                objaddon.objapplication.StatusBar.SetText("Item Count is mismatching with the Scrap Tab of Sub Contracting.Please check...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False : Exit Sub
                            End If
                        End If
                        For i As Integer = 1 To objmatrix.VisualRowCount
                            If objmatrix.Columns.Item("1").Cells.Item(i).Specific.String <> "" Then
                                If objaddon.HANA Then
                                    objRS.DoQuery("select ""U_Itemcode"",""U_WhsCode"",""U_Qty"",""U_Price"" from ""@MIPL_POR3"" where ""DocEntry""=" & objform.Items.Item("U_SubConNo").Specific.String & " and ""U_Status"" ='O'")
                                Else
                                    objRS.DoQuery("select U_Itemcode,U_WhsCode,U_Qty,U_Price from [@MIPL_POR3] where DocEntry=" & objform.Items.Item("U_SubConNo").Specific.String & " and U_Status ='O'")
                                End If
                                If objRS.RecordCount > 0 Then
                                    'Dim Value As String = objmatrix.Columns.Item("10").Cells.Item(i).Specific.String
                                    'Value = Value.Replace("INR ", "")
                                    If objmatrix.Columns.Item("1").Cells.Item(i).Specific.String <> objRS.Fields.Item("U_Itemcode").Value.ToString Or objmatrix.Columns.Item("15").Cells.Item(i).Specific.String <> objRS.Fields.Item("U_WhsCode").Value.ToString Or CDbl(objmatrix.Columns.Item("9").Cells.Item(i).Specific.String) <> CDbl(objRS.Fields.Item("U_Qty").Value.ToString) Then
                                        objaddon.objapplication.StatusBar.SetText("Line Item is not matching -Scrap Tab.Please re-generate from Sub Contracting screen.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False : Exit Sub
                                    End If
                                End If
                            End If
                        Next
                    End If
                    For i As Integer = 1 To objmatrix.VisualRowCount
                        If objmatrix.Columns.Item("U_TabType").Cells.Item(i).Specific.String = "" Then
                            If objmatrix.Columns.Item("1").Cells.Item(i).Specific.String <> "" Then
                                'If CDbl(objmatrix.Columns.Item("U_PlanQty").Cells.Item(i).Specific.String) <> CDbl(objmatrix.Columns.Item("9").Cells.Item(i).Specific.String) Then
                                '    objaddon.objapplication.SetStatusBarMessage("Quantity mismatching from Sub-Con Output Tab Line: " & i, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                '    BubbleEvent = False : Exit Sub
                                'End If
                                Dim Value As String = objmatrix.Columns.Item("10").Cells.Item(i).Specific.String
                                Value = Value.Replace("INR", "")
                                If CDbl(objmatrix.Columns.Item("U_Price").Cells.Item(i).Specific.String) <> CDbl(Value) Then
                                    objaddon.objapplication.SetStatusBarMessage("UnitPrice should not to change...", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                    BubbleEvent = False : Exit Sub
                                End If
                            End If
                        End If
                    Next
                    For j = 1 To objmatrix.VisualRowCount
                        If objmatrix.Columns.Item("U_TabType").Cells.Item(j).Specific.String = "Scrap" Then
                            If CDbl(objmatrix.Columns.Item("U_PlanQty").Cells.Item(j).Specific.String) <> CDbl(objmatrix.Columns.Item("9").Cells.Item(j).Specific.String) Then
                                objaddon.objapplication.SetStatusBarMessage("Quantity mismatching from Sub-Con Scrap Tab Line: " & j, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                BubbleEvent = False : Exit Sub
                            End If
                        End If
                    Next
                End If

                GC.Collect()
            Catch ex As Exception
                objaddon.objapplication.SetStatusBarMessage(ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try
        End Sub

        Private Sub Form_DataAddAfter(ByRef pVal As SAPbouiCOM.BusinessObjectInfo)
            Try
                If objform.Items.Item("U_SubConNo").Specific.String = "" Then Exit Sub
                Dim Status As String = "", TabName As String = "", StrQuery1 As String = "", StrQuery As String = "", GREntry As String = ""
                Dim objmatrix, objmatrixScrap, objSubConMatrix, objMatCosting As SAPbouiCOM.Matrix
                Dim objrs, objrecset As SAPbobsCOM.Recordset
                Dim FolderOutput, FolderScrap As SAPbouiCOM.Folder
                Try
                    If objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then Exit Sub
                    If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.ActionSuccess = True And pVal.BeforeAction = False Then
                        objmatrix = objform.Items.Item("13").Specific
                        objSubConform = objaddon.objapplication.Forms.GetForm("SUBCTPO", 1)
                        objSubConMatrix = objSubConform.Items.Item("mtxoutput").Specific
                        objmatrixScrap = objSubConform.Items.Item("mtxscrap").Specific
                        objMatCosting = objSubConform.Items.Item("MtxCosting").Specific
                        FolderOutput = objSubConform.Items.Item("flroutput").Specific
                        FolderScrap = objSubConform.Items.Item("flrscrap").Specific
                        objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        objrecset = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        If objform.Items.Item("U_SubConNo").Specific.String <> "" Then
                            If TranList = "Y" Then
                                GREntry = objform.DataSources.DBDataSources.Item("OIGN").GetValue("DocEntry", 0)
                            Else
                                GREntry = objform.DataSources.DBDataSources.Item("OIGN").GetValue("DocNum", 0)
                            End If
                            If objSubConform.Items.Item("txtGRNo").Specific.String = "" Then
                                objSubConform.Items.Item("txtGRNo").Specific.String = GREntry
                            Else
                                objSubConform.Items.Item("txtGRNo").Specific.String += ", " & GREntry
                            End If
                        End If
                        GREntry = objform.DataSources.DBDataSources.Item("OIGN").GetValue("DocEntry", 0)
                        Dim DocDate As Date = Date.ParseExact(objform.Items.Item("9").Specific.String, "dd/MM/yy", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                        If FolderOutput.Selected = True Then
                            For j = 1 To objSubConMatrix.VisualRowCount
                                If objSubConMatrix.Columns.Item("Code").Cells.Item(j).Specific.String <> "" And objSubConMatrix.Columns.Item("GRNo").Cells.Item(j).Specific.String = "" And objSubConMatrix.Columns.Item("Colstat").Cells.Item(j).Specific.String = "O" Then
                                    objSubConMatrix.Columns.Item("GRNo").Cells.Item(j).Specific.String = GREntry
                                    objSubConMatrix.Columns.Item("Colstat").Cells.Item(j).Specific.String = "C"
                                    If objaddon.HANA Then
                                        objSubConMatrix.Columns.Item("Instock").Cells.Item(j).Specific.String = objaddon.objglobalmethods.getSingleValue("select ""OnHand"" from  OITW where ""ItemCode""='" & objSubConMatrix.Columns.Item("Code").Cells.Item(j).Specific.String & "' and ""WhsCode""='" & objSubConMatrix.Columns.Item("Whse").Cells.Item(j).Specific.String & "'")
                                    Else
                                        objSubConMatrix.Columns.Item("Instock").Cells.Item(j).Specific.String = objaddon.objglobalmethods.getSingleValue("select OnHand from  OITW where ItemCode='" & objSubConMatrix.Columns.Item("Code").Cells.Item(j).Specific.String & "' and WhsCode='" & objSubConMatrix.Columns.Item("Whse").Cells.Item(j).Specific.String & "'")
                                    End If
                                    'objSubConMatrix.CommonSetting.SetRowEditable(j, False)
                                    objSubConMatrix.CommonSetting.SetCellEditable(j, 2, False)
                                    objSubConMatrix.CommonSetting.SetCellEditable(j, 5, False)
                                    objSubConMatrix.CommonSetting.SetCellEditable(j, 13, False)
                                    objSubConMatrix.CommonSetting.SetCellEditable(j, 20, False)
                                    objSubConMatrix.CommonSetting.SetCellEditable(j, 21, False)
                                End If
                            Next

                            'StrQuery1 = objaddon.objglobalmethods.getSingleValue("select To_Varchar(""DocDate"",'dd/MM/yy') ""RefDate"" from OIGE where ""DocEntry""='" & objSubConMatrix.Columns.Item("GINo").Cells.Item(objSubConMatrix.VisualRowCount).Specific.String & "' ")
                            'Dim RefDate As Date = Date.ParseExact(StrQuery1, "dd/MM/yy", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                            'StrQuery = "update T0 set T0.""U_Status""='C' from OJDT T0 join ""@MIPL_OPOR"" T1 on T1.""DocEntry""=T0.""U_SubConNo"" join ""@MIPL_POR5"" T2 on T2.""DocEntry""=T1.""DocEntry"""
                            'StrQuery += vbCrLf + " where T0.""U_Status""='O' and T0.""U_SubConNo""='" & objSubConform.Items.Item("txtentry").Specific.String & "' and T0.""RefDate""='" & RefDate.ToString("yyyyMMdd") & "' and T2.""U_Status""='O'"
                            'objrecset.DoQuery(StrQuery)
                            'For j = 1 To objMatCosting.VisualRowCount
                            '    If objMatCosting.Columns.Item("Code").Cells.Item(j).Specific.String <> "" And objMatCosting.Columns.Item("Stat").Cells.Item(j).Specific.String = "O" Then
                            '        objMatCosting.Columns.Item("Stat").Cells.Item(j).Specific.String = "C"
                            '        objMatCosting.CommonSetting.SetRowEditable(j, False)
                            '    End If
                            'Next
                        End If
                        If FolderScrap.Selected = True Then
                            For j = 1 To objmatrixScrap.VisualRowCount
                                If objmatrixScrap.Columns.Item("Code").Cells.Item(j).Specific.String <> "" And objmatrixScrap.Columns.Item("Colstat").Cells.Item(j).Specific.String = "O" And objmatrixScrap.Columns.Item("GRNo").Cells.Item(j).Specific.String = "" Then
                                    objmatrixScrap.Columns.Item("GRNo").Cells.Item(j).Specific.String = GREntry
                                    objmatrixScrap.Columns.Item("Colstat").Cells.Item(j).Specific.String = "C"
                                    If objaddon.HANA Then
                                        objmatrixScrap.Columns.Item("Instock").Cells.Item(j).Specific.String = objaddon.objglobalmethods.getSingleValue("select ""OnHand"" from  OITW where ""ItemCode""='" & objmatrixScrap.Columns.Item("Code").Cells.Item(j).Specific.String & "' and ""WhsCode""='" & objmatrixScrap.Columns.Item("Whse").Cells.Item(j).Specific.String & "'")
                                    Else
                                        objmatrixScrap.Columns.Item("Instock").Cells.Item(j).Specific.String = objaddon.objglobalmethods.getSingleValue("select OnHand from  OITW where ItemCode='" & objmatrixScrap.Columns.Item("Code").Cells.Item(j).Specific.String & "' and WhsCode='" & objmatrixScrap.Columns.Item("Whse").Cells.Item(j).Specific.String & "'")
                                    End If
                                    ' objmatrixScrap.CommonSetting.SetRowEditable(j, False)
                                    objmatrixScrap.CommonSetting.SetCellEditable(j, 2, False)
                                    objmatrixScrap.CommonSetting.SetCellEditable(j, 3, False)
                                    objmatrixScrap.CommonSetting.SetCellEditable(j, 5, False)
                                    objmatrixScrap.CommonSetting.SetCellEditable(j, 6, False)
                                    objmatrixScrap.CommonSetting.SetCellEditable(j, 9, False)
                                    objmatrixScrap.CommonSetting.SetCellEditable(j, 16, False)
                                    objmatrixScrap.CommonSetting.SetCellEditable(j, 17, False)
                                End If
                            Next
                        End If
                        'If objSubConform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objSubConform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        'If objSubConform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        '    objSubConform.Items.Item("1").Click()
                        '    Exit Sub
                        'End If
                        'objSubConform.Refresh()
                        'objSubConform.Update()
                        If objSubConform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objSubConform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        If objSubConform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                            objSubConform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        End If
                        'objSubConform.Update()
                    End If
                    objrs = Nothing
                    objrecset = Nothing
                    GC.Collect()
                Catch ex As Exception
                    objaddon.objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                End Try
            Catch ex As Exception

            End Try
        End Sub

        Private Sub Form_DataLoadAfter(ByRef pVal As SAPbouiCOM.BusinessObjectInfo)
            Try
                If objform.Items.Item("U_SubConNo").Specific.String = "" Then Exit Sub
                objform = objaddon.objapplication.Forms.GetForm("721", pVal.FormTypeEx)
                Dim oUDFForm As SAPbouiCOM.Form
                oUDFForm = objaddon.objapplication.Forms.Item(objform.UDFFormUID)
                oUDFForm.Items.Item("U_SubConNo").Enabled = False
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Form_DeactivateAfter(pVal As SAPbouiCOM.SBOItemEventArg)
            Try
                'If Not objaddon.objapplication.Forms.ActiveForm.TypeEx = "41" And Not objaddon.objapplication.Forms.ActiveForm.TypeEx = "21" Then  'Batch
                If objaddon.objapplication.Forms.ActiveForm.TypeEx = "SUBCTPO" Then
                    objSubConform = objaddon.objapplication.Forms.GetForm("SUBCTPO", 1)
                    If objSubConform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objSubConform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    If objSubConform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        objSubConform.Items.Item("1").Click()
                    End If
                End If

            Catch ex As Exception

            End Try

        End Sub

        Private Sub Form_LoadAfter(pVal As SAPbouiCOM.SBOItemEventArg)
            Try
                objform = objaddon.objapplication.Forms.GetForm("721", pVal.FormTypeCount)

            Catch ex As Exception

            End Try


        End Sub
    End Class
End Namespace
