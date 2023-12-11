Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework

Namespace SubContractingPO
    <FormAttribute("940", "SubContractingPO/FrmInventoryTransfer.b1f")>
    Friend Class FrmInventoryTransfer
        Inherits SystemFormBase
        Private WithEvents objSubConform As SAPbouiCOM.Form
        Private WithEvents objform As SAPbouiCOM.Form
        Public WithEvents objmatrix As SAPbouiCOM.Matrix
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
            AddHandler LoadAfter, AddressOf Me.Form_LoadAfter
            AddHandler DeactivateAfter, AddressOf Me.Form_DeactivateAfter
            AddHandler DataLoadAfter, AddressOf Me.Form_DataLoadAfter

        End Sub

        Private WithEvents StaticText0 As SAPbouiCOM.StaticText

        Private Sub OnCustomInitialize()
            Try
                objform = objaddon.objapplication.Forms.GetForm("940", 0)
                objSubConform = objaddon.objapplication.Forms.GetForm("SUBCTPO", 0)

            Catch ex As Exception
            End Try

        End Sub

        Private WithEvents EditText0 As SAPbouiCOM.EditText

        Private Sub Form_DataAddAfter(ByRef pVal As SAPbouiCOM.BusinessObjectInfo)
            If objform.Items.Item("U_SubConNo").Specific.String = "" Then Exit Sub
            Dim Status As String = ""
            Dim Docentry As String
            If objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then Exit Sub
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.ActionSuccess = True And pVal.BeforeAction = False Then
                    If TranList = "Y" Then
                        Docentry = objform.DataSources.DBDataSources.Item("OWTR").GetValue("DocEntry", 0)
                    Else
                        Docentry = objform.DataSources.DBDataSources.Item("OWTR").GetValue("DocNum", 0)
                    End If
                    objmatrix = objform.Items.Item("23").Specific
                    objSubConform = objaddon.objapplication.Forms.GetForm("SUBCTPO", 0)
                    If objform.Items.Item("U_SubConNo").Specific.String <> "" Then
                        If objSubConform.Items.Item("TxtInvTr").Specific.String = "" Then
                            objSubConform.Items.Item("TxtInvTr").Specific.String = Docentry
                        Else
                            objSubConform.Items.Item("TxtInvTr").Specific.String += ", " & Docentry
                        End If
                    End If
                    'If objSubConform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objSubConform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    '    objSubConform.Items.Item("1").Click()
                    'End If
                End If
                GC.Collect()
            Catch ex As Exception
            End Try

        End Sub

        Private Sub Form_LoadAfter(pVal As SAPbouiCOM.SBOItemEventArg)
            Try
                StaticText0.Item.Top = objform.Items.Item("37").Top + 15 '502
                StaticText0.Item.Left = objform.Items.Item("37").Left  '150
                EditText0.Item.Top = objform.Items.Item("36").Top + 15 '19
                EditText0.Item.Left = objform.Items.Item("36").Left '19

            Catch ex As Exception

            End Try

        End Sub

        Private Sub Form_DeactivateAfter(pVal As SAPbouiCOM.SBOItemEventArg)
            Try
                ' If Not objaddon.objapplication.Forms.ActiveForm.TypeEx = "42" And Not objaddon.objapplication.Forms.ActiveForm.TypeEx = "21" Then  'Batch
                If objaddon.objapplication.Forms.ActiveForm.TypeEx = "SUBCTPO" Then
                        objSubConform = objaddon.objapplication.Forms.GetForm("SUBCTPO", 0)
                        If objSubConform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objSubConform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        If objSubConform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                            objSubConform.Items.Item("1").Click()
                        End If
                    End If

                'End If

            Catch ex As Exception

            End Try

        End Sub

        Private Sub Form_DataLoadAfter(ByRef pVal As SAPbouiCOM.BusinessObjectInfo)
            Try
                If objform.Items.Item("U_SubConNo").Specific.String = "" Then Exit Sub
                objform = objaddon.objapplication.Forms.GetForm("940", 0)
                Dim oUDFForm As SAPbouiCOM.Form
                oUDFForm = objaddon.objapplication.Forms.Item(objform.UDFFormUID)
                oUDFForm.Items.Item("U_SubConNo").Enabled = False
                oUDFForm.Items.Item("U_ScrapNum").Enabled = False
            Catch ex As Exception

            End Try

        End Sub

        Private WithEvents Button0 As SAPbouiCOM.Button

        Private Sub Button0_ClickBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Button0.ClickBefore
            Try
                If objform.Items.Item("U_SubConNo").Specific.String = "" Then Exit Sub
                Dim objRS As SAPbobsCOM.Recordset
                Dim StrQuery, LineId As String
                objRS = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    objmatrix = objform.Items.Item("23").Specific
                    Dim Rowcount As Integer = objmatrix.VisualRowCount - 1
                    If objform.Items.Item("U_SubConNo").Specific.String <> "" Then
                        If objaddon.HANA Then
                            StrQuery = "select ""U_Itemcode"",""LineId"" from ""@MIPL_POR1"" where ""DocEntry""=" & objform.Items.Item("U_SubConNo").Specific.String & ""
                        Else
                            StrQuery = "select U_Itemcode,LineId from [@MIPL_POR1] where DocEntry=" & objform.Items.Item("U_SubConNo").Specific.String & ""
                        End If
                        objRS.DoQuery(StrQuery)
                    End If
                    If objRS.RecordCount > 0 And Rowcount > 0 Then
                        If Rowcount <> objRS.RecordCount Then
                            objaddon.objapplication.StatusBar.SetText("Item Count is mismatching with the Input Tab of Sub Contracting.Please check...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False : Exit Sub
                        End If
                    End If
                    For i As Integer = 1 To objmatrix.VisualRowCount
                        If objmatrix.Columns.Item("1").Cells.Item(i).Specific.String <> "" Then
                            If objmatrix.Columns.Item("U_LineID").Cells.Item(i).Specific.String = "" Then
                                LineId = "0"
                            Else
                                LineId = objmatrix.Columns.Item("U_LineID").Cells.Item(i).Specific.String
                            End If
                            If objaddon.HANA Then
                                objRS.DoQuery("select ""U_Itemcode"",""U_WhsCode"",""U_SubWhse"" from ""@MIPL_POR1"" where ""DocEntry""=" & objform.Items.Item("U_SubConNo").Specific.String & " and ""LineId"" ='" & LineId & "'")
                            Else
                                objRS.DoQuery("select U_Itemcode,U_WhsCode,U_SubWhse from [@MIPL_POR1] where DocEntry=" & objform.Items.Item("U_SubConNo").Specific.String & " and LineId ='" & LineId & "'")
                            End If
                            If objRS.RecordCount > 0 Then
                                If objmatrix.Columns.Item("1").Cells.Item(i).Specific.String <> objRS.Fields.Item("U_Itemcode").Value.ToString Or objmatrix.Columns.Item("1470001039").Cells.Item(i).Specific.String <> objRS.Fields.Item("U_WhsCode").Value.ToString Or objmatrix.Columns.Item("5").Cells.Item(i).Specific.String <> objRS.Fields.Item("U_SubWhse").Value.ToString Then
                                    objaddon.objapplication.StatusBar.SetText("Line Item Details is not matching.Please re-generate from Sub Contracting screen.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False : Exit Sub
                                End If
                            End If
                        End If
                    Next
                    For j = 1 To objmatrix.VisualRowCount
                        If objmatrix.Columns.Item("1").Cells.Item(j).Specific.String <> "" Then
                            If CDbl(objmatrix.Columns.Item("U_PlanQty").Cells.Item(j).Specific.String) < CDbl(objmatrix.Columns.Item("10").Cells.Item(j).Specific.String) Then
                                objaddon.objapplication.SetStatusBarMessage("Quantity mismatching from Sub-Con Input Tab Line: " & j, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                BubbleEvent = False : Exit Sub
                            End If
                        End If
                    Next
                End If
                objRS = Nothing
                GC.Collect()
            Catch ex As Exception

            End Try

        End Sub


    End Class
End Namespace
