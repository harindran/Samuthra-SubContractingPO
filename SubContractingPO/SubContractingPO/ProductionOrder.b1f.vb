Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework

Namespace SubContractingPO
    <FormAttribute("65211", "SubContractingPO/ProductionOrder.b1f")>
    Friend Class ProductionOrder
        Inherits SystemFormBase
        Private WithEvents objSubConform As SAPbouiCOM.Form
        Private WithEvents objform As SAPbouiCOM.Form
        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.Button0 = CType(Me.GetItem("1").Specific, SAPbouiCOM.Button)
            Me.Button1 = CType(Me.GetItem("BtnSubCon").Specific, SAPbouiCOM.Button)
            Me.Button2 = CType(Me.GetItem("2").Specific, SAPbouiCOM.Button)
            Me.EditText0 = CType(Me.GetItem("6").Specific, SAPbouiCOM.EditText)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()
            AddHandler DataLoadAfter, AddressOf Me.Form_DataLoadAfter
            AddHandler LoadAfter, AddressOf Me.Form_LoadAfter

        End Sub
        Private WithEvents Button0 As SAPbouiCOM.Button

        Private Sub OnCustomInitialize()

            Try
                objform = objaddon.objapplication.Forms.GetForm("65211", 0)
                'objSubConform = objaddon.objapplication.Forms.GetForm("SUBCTPO", 0)
                If objaddon.HANA Then
                    SubPOScreen = objaddon.objglobalmethods.getSingleValue("select ""U_SubScreen"" from ""@MIPL_GEN"" where ifnull(""U_SubScreen"",'')='Y'")
                Else
                    SubPOScreen = objaddon.objglobalmethods.getSingleValue("select U_SubScreen from [@MIPL_GEN] where isnull(U_SubScreen,'')='Y'")
                End If
                If SubPOScreen = "Y" Then
                    Button1.Item.Visible = False
                End If

            Catch ex As Exception

            End Try
        End Sub
        Private Sub Button0_ClickBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles Button0.ClickBefore
            'Dim objMatrix As SAPbouiCOM.Matrix
            'Dim Stock As Double = 0
            'Try
            '    objMatrix = objform.Items.Item("37").Specific
            '    For i As Integer = 1 To objMatrix.VisualRowCount
            '        If objMatrix.Columns.Item("4").Cells.Item(i).Specific.string <> "" Then
            '            Stock = objaddon.objglobalmethods.getSingleValue("select Sum(""OnHand"") from OITW where ""ItemCode""='" & objMatrix.Columns.Item("4").Cells.Item(i).Specific.string & "' and ""WhsCode""='" & objMatrix.Columns.Item("10").Cells.Item(i).Specific.string & "' ")
            '            If Stock < CDbl(objMatrix.Columns.Item("14").Cells.Item(i).Specific.string) Then
            '                objaddon.objapplication.SetStatusBarMessage("Please create a Sub contracting for this Production Item...", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            '                'BubbleEvent = False : Exit Sub
            '            End If
            '        End If
            '    Next
            'Catch ex As Exception

            'End Try

        End Sub

        Public Sub CreateButton()

            Dim objButton As SAPbouiCOM.Button
            Dim objItem As SAPbouiCOM.Item
            Try
                objform = objaddon.objapplication.Forms.GetForm("65211", 0)
                objItem = objform.Items.Add("BtnAdd", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
                objItem.Left = objform.Items.Item("2").Left + objform.Items.Item("2").Width + 10
                objItem.Width = 100
                objItem.Top = objform.Items.Item("2").Top
                objItem.Height = objform.Items.Item("2").Height
                objButton = objItem.Specific
                objButton.Caption = "Sub Contractor"
                objaddon.objapplication.SetStatusBarMessage("Button Created", SAPbouiCOM.BoMessageTime.bmt_Short, False)
            Catch ex As Exception
            End Try

        End Sub
        
        Private WithEvents Button1 As SAPbouiCOM.Button

        Private Sub Button1_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button1.ClickAfter
            Try
                Dim DocEntry, PONum, POEntry As String, StrQuery As String = "", CheckBOMItem As String = ""
                Dim objrs As SAPbobsCOM.Recordset
                Dim DocDate As Date
                objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    If Button1.Item.Enabled = True Then
                        DocDate = Date.ParseExact(objform.Items.Item("24").Specific.String, "dd/MM/yy", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                        If objaddon.HANA Then
                            DocEntry = objaddon.objglobalmethods.getSingleValue("select Top 1 ""DocEntry"" from OWOR where ""DocNum""='" & objform.Items.Item("18").Specific.String & "' and ""PostDate""='" & DocDate.ToString("yyyMMdd") & "' ")
                            PONum = objaddon.objglobalmethods.getSingleValue("select ""DocEntry"" from ""@MIPL_OPOR"" where ""U_PONum""='" & DocEntry & "'")
                        Else
                            DocEntry = objaddon.objglobalmethods.getSingleValue("select Top 1 DocEntry from OWOR where DocNum='" & objform.Items.Item("18").Specific.String & "' and PostDate='" & DocDate.ToString("yyyMMdd") & "' ")
                            PONum = objaddon.objglobalmethods.getSingleValue("select DocEntry from [@MIPL_OPOR] where U_PONum='" & DocEntry & "'")
                        End If
                        Link_Value = "-1"
                        If PONum <> "" Then
                            Link_Value = PONum 'DocEntry
                            Link_objtype = "202"
                            Dim activeform As New SubContractingPO
                            activeform.Show()
                        Else
                            Dim activeform As New SubContractingPO
                            activeform.Show()
                            POEntry = objform.DataSources.DBDataSources.Item("OWOR").GetValue("DocEntry", 0)
                            activeform.objform.Items.Item("txtponum").Specific.String = POEntry
                            If objaddon.HANA Then
                                StrQuery = "Select ""ItemCode"",""ProdName"",""PlannedQty"" from OWOR where ""DocEntry""='" & POEntry & "'"
                            Else
                                StrQuery = "Select ItemCode,ProdName,PlannedQty from OWOR where DocEntry='" & POEntry & "'"
                            End If
                            objrs.DoQuery(StrQuery)
                            If objrs.RecordCount > 0 Then
                                If objaddon.HANA Then
                                    CheckBOMItem = objaddon.objglobalmethods.getSingleValue("select 1 from ""@MIPL_BOM1"" where ""U_ItemCode""='" & objrs.Fields.Item("ItemCode").Value.ToString & "'")
                                Else
                                    CheckBOMItem = objaddon.objglobalmethods.getSingleValue("select 1 from [@MIPL_BOM1] where U_ItemCode='" & objrs.Fields.Item("ItemCode").Value.ToString & "'")
                                End If
                                If CheckBOMItem <> "" Then
                                    activeform.objform.Items.Item("txtsitem").Specific.String = objrs.Fields.Item("ItemCode").Value.ToString
                                    activeform.objform.Items.Item("SItemDesc").Specific.String = objrs.Fields.Item("ProdName").Value.ToString
                                    activeform.objform.Items.Item("SQty").Specific.String = objrs.Fields.Item("PlannedQty").Value.ToString
                                Else
                                    objaddon.objapplication.StatusBar.SetText("Sub-BOM Not Created for this ItemCode: " & objrs.Fields.Item("ItemCode").Value.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                End If
                            End If
                            'objSubConform.Items.Item("txtponum").Specific.String = POEntry
                        End If
                    End If
                End If
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub

        Private WithEvents Button2 As SAPbouiCOM.Button
        Private WithEvents EditText0 As SAPbouiCOM.EditText

        Private Sub Form_DataLoadAfter(ByRef pVal As SAPbouiCOM.BusinessObjectInfo)
            Try
                objaddon.objglobalmethods.SubConButtonEnabling(objform)
            Catch ex As Exception

            End Try
        End Sub

        Private Sub Form_LoadAfter(pVal As SAPbouiCOM.SBOItemEventArg)
            Try
                Button1.Item.Top = objform.Items.Item("2").Top '502
                Button1.Item.Left = objform.Items.Item("2").Left + 85 '150
                Button1.Item.Height = objform.Items.Item("2").Height '19
            Catch ex As Exception

            End Try

        End Sub
    End Class
End Namespace
