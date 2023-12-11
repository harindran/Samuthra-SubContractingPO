Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework

Namespace SubContractingPO
    <FormAttribute("672", "SubContractingPO/FrmSAP_BOM.b1f")>
    Friend Class FrmSAP_BOM
        Inherits SystemFormBase
        Private WithEvents objform As SAPbouiCOM.Form
        Private WithEvents objMatrix As SAPbouiCOM.Matrix
        Private Shared FormCount As Integer = 0
        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.Button0 = CType(Me.GetItem("1").Specific, SAPbouiCOM.Button)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()
            AddHandler DataAddAfter, AddressOf Me.Form_DataAddAfter
            AddHandler DataUpdateAfter, AddressOf Me.Form_DataUpdateAfter
            AddHandler CloseAfter, AddressOf Me.Form_CloseAfter

        End Sub
        Private WithEvents Button0 As SAPbouiCOM.Button

        Private Sub OnCustomInitialize()
            Try
                FormCount += 1
                objform = objaddon.objapplication.Forms.GetForm("672", FormCount)
                objMatrix = objform.Items.Item("3").Specific
            Catch ex As Exception

            End Try
        End Sub

        Private Function Creating_SubBOM_From_SAPBOM() As Boolean
            Try
                Dim GetSubConItem, getCopytoBOM, InvntItem As String
                Dim Price, Total, currency As String
                Dim Flag As Boolean = False
                Dim objType, Whse As SAPbouiCOM.ComboBox
                Dim oGeneralService As SAPbobsCOM.GeneralService
                Dim oGeneralData As SAPbobsCOM.GeneralData
                Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
                Dim oGeneralDataCollection As SAPbobsCOM.GeneralDataCollection
                Dim oChild As SAPbobsCOM.GeneralData
                If objaddon.HANA Then
                    getCopytoBOM = objaddon.objglobalmethods.getSingleValue("select ""U_CopBOM"" from ""@MIPL_GEN"" where ifnull(""U_CopBOM"",'')='Y'")
                Else
                    getCopytoBOM = objaddon.objglobalmethods.getSingleValue("select U_CopBOM from [@MIPL_GEN] where isnull(U_CopBOM,'')='Y'")
                End If

                If getCopytoBOM = "Y" Then
                    If objaddon.HANA Then
                        GetSubConItem = objaddon.objglobalmethods.getSingleValue("Select ""U_SubCont"" from OITM where ""ItemCode"" ='" & objform.Items.Item("4").Specific.String & "'")
                    Else
                        GetSubConItem = objaddon.objglobalmethods.getSingleValue("Select U_SubCont from OITM where ItemCode ='" & objform.Items.Item("4").Specific.String & "'")
                    End If
                    If GetSubConItem = "Y" Then
                        oGeneralService = objaddon.objcompany.GetCompanyService.GetGeneralService("SUBBOM")
                        oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                        oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                        oGeneralDataCollection = oGeneralData.Child("MIPL_BOM1")
                        Try
                            'If objaddon.HANA Then
                            '    getCopytoBOM = objaddon.objglobalmethods.getSingleValue("Select max(""U_Type"")+1 ""Count"" from ""@MIPL_OBOM"" where ""U_ItemCode""='" & objform.Items.Item("4").Specific.String & "'")
                            'Else
                            '    getCopytoBOM = objaddon.objglobalmethods.getSingleValue("Select max(U_Type)+1 Count from [@MIPL_OBOM] where U_ItemCode='" & objform.Items.Item("4").Specific.String & "'")
                            'End If
                            'oGeneralParams.SetProperty("U_ItemCode", objform.Items.Item("4").Specific.String)
                            oGeneralParams.SetProperty("Code", objform.Items.Item("4").Specific.String)
                            oGeneralData = oGeneralService.GetByParams(oGeneralParams)
                            Flag = True
                        Catch ex As Exception
                            Flag = False
                        End Try
                        If Flag = True Then
                            If objaddon.objapplication.MessageBox("This BOM is already there in Sub-BOM. Do you want to update the Sub-BOM?", 2, "Yes", "No") <> 1 Then Return True
                            objaddon.objapplication.StatusBar.SetText("Modifying SAP BOM to Sub-Contracting BOM.Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        Else
                            objaddon.objapplication.StatusBar.SetText("Copying SAP BOM to Sub-Contracting BOM.Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        End If
                        oGeneralData.SetProperty("Code", objform.Items.Item("4").Specific.String)
                        oGeneralData.SetProperty("Name", objform.Items.Item("5").Specific.String)
                        oGeneralData.SetProperty("U_ItemCode", objform.Items.Item("4").Specific.String)
                        Whse = objform.Items.Item("16").Specific
                        oGeneralData.SetProperty("U_BOMType", "1")
                        oGeneralData.SetProperty("U_WhseCode", Whse.Selected.Value)
                        oGeneralData.SetProperty("U_Qty", CDbl(objform.Items.Item("17").Specific.String))
                        If objaddon.HANA Then
                            currency = objaddon.objglobalmethods.getSingleValue("Select ""MainCurncy"" from OADM")
                        Else
                            currency = objaddon.objglobalmethods.getSingleValue("Select MainCurncy from OADM")
                        End If
                        oChild = oGeneralDataCollection.Add()
                        Try
                            Dim RowCount As Integer
                            RowCount = oGeneralData.Child("MIPL_BOM1").Count - 1
                            'MatrixCount = objMatrix.VisualRowCount
                            If RowCount >= 0 Then
                                For deleterow As Integer = RowCount To 0 Step -1
                                    oGeneralData.Child("MIPL_BOM1").Remove(deleterow)
                                Next
                            End If
                        Catch ex As Exception
                            Return False
                        End Try
                        For i As Integer = 1 To objMatrix.VisualRowCount
                            If objMatrix.Columns.Item("1").Cells.Item(i).Specific.String <> "" Then
                                If objaddon.HANA Then
                                    InvntItem = objaddon.objglobalmethods.getSingleValue("select ""InvntItem"" from OITM where ""ItemCode""='" & objMatrix.Columns.Item("1").Cells.Item(i).Specific.String & "'")
                                Else
                                    InvntItem = objaddon.objglobalmethods.getSingleValue("select InvntItem from OITM where ItemCode='" & objMatrix.Columns.Item("1").Cells.Item(i).Specific.String & "'")
                                End If
                                If InvntItem = "Y" Then
                                    objType = objMatrix.Columns.Item("1880000002").Cells.Item(i).Specific
                                    Price = objMatrix.Columns.Item("4").Cells.Item(i).Specific.String
                                    Price = Trim(Price.Replace(currency, "")) 'System.Text.RegularExpressions.Regex.Replace(Price, "[^\d]", "")
                                    Total = objMatrix.Columns.Item("5").Cells.Item(i).Specific.String
                                    Total = Trim(Total.Replace(currency, "")) 'System.Text.RegularExpressions.Regex.Replace(Total, "[^\d]", "")
                                    If i > oGeneralData.Child("MIPL_BOM1").Count Then
                                        oGeneralData.Child("MIPL_BOM1").Add()
                                    End If

                                    If objType.Selected.Value = "4" Then
                                        'oGeneralData.Child("MIPL_BOM1").Item(i - 1).SetProperty("LineId", i - 1)
                                        oGeneralData.Child("MIPL_BOM1").Item(i - 1).SetProperty("U_Type", "4")
                                        oGeneralData.Child("MIPL_BOM1").Item(i - 1).SetProperty("U_Itemcode", objMatrix.Columns.Item("1").Cells.Item(i).Specific.String)
                                        oGeneralData.Child("MIPL_BOM1").Item(i - 1).SetProperty("U_ItemDesc", objMatrix.Columns.Item("44").Cells.Item(i).Specific.String)
                                        oGeneralData.Child("MIPL_BOM1").Item(i - 1).SetProperty("U_Qty", CDbl(objMatrix.Columns.Item("2").Cells.Item(i).Specific.String))
                                        oGeneralData.Child("MIPL_BOM1").Item(i - 1).SetProperty("U_UOMName", objMatrix.Columns.Item("45").Cells.Item(i).Specific.String)
                                        oGeneralData.Child("MIPL_BOM1").Item(i - 1).SetProperty("U_Whse", objMatrix.Columns.Item("3").Cells.Item(i).Specific.String)
                                        oGeneralData.Child("MIPL_BOM1").Item(i - 1).SetProperty("U_Unitprice", Price)
                                        oGeneralData.Child("MIPL_BOM1").Item(i - 1).SetProperty("U_Total", Total)
                                        oGeneralData.Child("MIPL_BOM1").Item(i - 1).SetProperty("U_Comments", objMatrix.Columns.Item("48").Cells.Item(i).Specific.String)
                                        'oGeneralData.Child("MIPL_BOM1").Add()
                                        'oChild.SetProperty("U_Type", "4")
                                        'oChild.SetProperty("U_Itemcode", objMatrix.Columns.Item("1").Cells.Item(i).Specific.String)
                                        'oChild.SetProperty("U_ItemDesc", objMatrix.Columns.Item("44").Cells.Item(i).Specific.String)
                                        'oChild.SetProperty("U_Qty", CDbl(objMatrix.Columns.Item("2").Cells.Item(i).Specific.String))
                                        'oChild.SetProperty("U_UOMName", objMatrix.Columns.Item("45").Cells.Item(i).Specific.String)
                                        'oChild.SetProperty("U_Whse", objMatrix.Columns.Item("3").Cells.Item(i).Specific.String)
                                        'oChild.SetProperty("U_Unitprice", Price)
                                        'oChild.SetProperty("U_Total", Total)
                                        'oChild.SetProperty("U_Comments", objMatrix.Columns.Item("48").Cells.Item(i).Specific.String)
                                    ElseIf objType.Selected.Value = "290" Then
                                        oGeneralData.Child("MIPL_BOM1").Item(i - 1).SetProperty("U_Type", "290")
                                        oGeneralData.Child("MIPL_BOM1").Item(i - 1).SetProperty("U_Itemcode", objMatrix.Columns.Item("1").Cells.Item(i).Specific.String)
                                        oGeneralData.Child("MIPL_BOM1").Item(i - 1).SetProperty("U_ItemDesc", objMatrix.Columns.Item("44").Cells.Item(i).Specific.String)
                                        oGeneralData.Child("MIPL_BOM1").Item(i - 1).SetProperty("U_Qty", CDbl(objMatrix.Columns.Item("2").Cells.Item(i).Specific.String))
                                        oGeneralData.Child("MIPL_BOM1").Item(i - 1).SetProperty("U_UOMName", objMatrix.Columns.Item("45").Cells.Item(i).Specific.String)
                                        oGeneralData.Child("MIPL_BOM1").Item(i - 1).SetProperty("U_Whse", objMatrix.Columns.Item("3").Cells.Item(i).Specific.String)
                                        oGeneralData.Child("MIPL_BOM1").Item(i - 1).SetProperty("U_Unitprice", Price)
                                        oGeneralData.Child("MIPL_BOM1").Item(i - 1).SetProperty("U_Total", Total)
                                        oGeneralData.Child("MIPL_BOM1").Item(i - 1).SetProperty("U_Comments", objMatrix.Columns.Item("48").Cells.Item(i).Specific.String)
                                    End If
                                End If
                            End If
                        Next
                        If Flag = True Then
                            oGeneralService.Update(oGeneralData)
                            objaddon.objapplication.StatusBar.SetText("Modified Successfully from SAP BOM to Sub-Contracting BOM...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            Return True
                        Else
                            oGeneralParams = oGeneralService.Add(oGeneralData)
                            objaddon.objapplication.StatusBar.SetText("Copied Successfully from SAP BOM to Sub-Contracting BOM...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            Return True
                        End If
                    Else
                        Return True
                    End If
                Else
                    Return True
                End If


            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End Try
        End Function

        Private Sub Form_DataAddAfter(ByRef pVal As SAPbouiCOM.BusinessObjectInfo)
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then Exit Sub
                If Creating_SubBOM_From_SAPBOM() = False Then
                    objaddon.objapplication.StatusBar.SetText("Sub Contracting BOM has not been created. Please update either from SAP BOM or create manually from Sub-BOM...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            Catch ex As Exception
            End Try
        End Sub

        Private Sub Form_DataUpdateAfter(ByRef pVal As SAPbouiCOM.BusinessObjectInfo)
            Try
                'objform = objaddon.objapplication.Forms.GetForm("672", 0)
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then Exit Sub
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    If Creating_SubBOM_From_SAPBOM() = False Then
                        objaddon.objapplication.StatusBar.SetText("Sub Contracting BOM has not been updated. Please update manually from Sub-BOM...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End If
                End If
            Catch ex As Exception
            End Try
        End Sub

        Private Sub Form_CloseAfter(pVal As SAPbouiCOM.SBOItemEventArg)
            Try
                FormCount -= 1
            Catch ex As Exception
            End Try
        End Sub
    End Class
End Namespace
