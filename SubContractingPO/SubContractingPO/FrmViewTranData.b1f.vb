Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework

Namespace SubContractingPO
    <FormAttribute("TRANVIEW", "SubContractingPO/FrmViewTranData.b1f")>
    Friend Class FrmViewTranData
        Inherits UserFormBase
        Public WithEvents objform As SAPbouiCOM.Form
        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.Button0 = CType(Me.GetItem("101").Specific, SAPbouiCOM.Button)
            Me.Button1 = CType(Me.GetItem("2").Specific, SAPbouiCOM.Button)
            Me.Grid0 = CType(Me.GetItem("gridData").Specific, SAPbouiCOM.Grid)
            Me.Button2 = CType(Me.GetItem("Item_3").Specific, SAPbouiCOM.Button)
            Me.Button3 = CType(Me.GetItem("Item_4").Specific, SAPbouiCOM.Button)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()

        End Sub
        Private WithEvents Button0 As SAPbouiCOM.Button

        Private Sub OnCustomInitialize()
            Try
                objform = objaddon.objapplication.Forms.GetForm("TRANVIEW", 0)
                objform = objaddon.objapplication.Forms.ActiveForm
            Catch ex As Exception

            End Try
        End Sub
        Private WithEvents Button1 As SAPbouiCOM.Button
        Private WithEvents Grid0 As SAPbouiCOM.Grid

        Public Sub ViewTransactionData(ByVal Header As String, ByVal Line As String, ByVal LinkedID As String, ByVal InvTransfer As String, ByVal DocEntry As String)
            Dim objrs As SAPbobsCOM.Recordset
            Try
                Dim str_sql As String = ""
                If AutoProduction = "Y" Then
                    If InvTransfer = "Y" Then
                        If objaddon.HANA Then
                            str_sql = "select Distinct T0.""DocEntry"",T0.""DocNum"",T0.""DocDate"",T1.""ItemCode"",T1.""Dscription"",T1.""Quantity"" ,T0.""Comments"""
                            str_sql += vbCrLf + "from " & Header & " T0 join " & Line & " T1 on T0.""DocEntry""=T1.""DocEntry""  where T0.""U_SubConNo""='" & DocEntry & "';"
                        Else
                            str_sql = "select Distinct T0.DocEntry,T0.DocNum,T0.DocDate,T1.ItemCode,T1.Dscription,T1.Quantity ,T0.Comments"
                            str_sql += vbCrLf + "from " & Header & " T0 join " & Line & " T1 on T0.DocEntry=T1.DocEntry  where T0.U_SubConNo='" & DocEntry & "'"
                        End If
                    Else
                        If objaddon.HANA Then
                            str_sql = "select Distinct T0.""DocEntry"",T0.""DocNum"",T0.""DocDate"",T1.""ItemCode"",T1.""Dscription"",T1.""Quantity"" ,T0.""Comments"""
                            str_sql += vbCrLf + "from " & Header & " T0 join " & Line & " T1 on T0.""DocEntry""=T1.""DocEntry""  where T1.""BaseType""=202  and T1.""BaseEntry""='" & DocEntry & "';"
                        Else
                            str_sql = "select Distinct T0.DocEntry,T0.DocNum,T0.DocDate,T1.ItemCode,T1.Dscription,T1.Quantity ,T0.Comments"
                            str_sql += vbCrLf + "from " & Header & " T0 join " & Line & " T1 on T0.DocEntry=T1.DocEntry  where T1.BaseType=202  and T1.BaseEntry='" & DocEntry & "'"
                        End If
                    End If
                End If
                objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objrs.DoQuery(str_sql)
                If objrs.RecordCount = 0 Then objaddon.objapplication.StatusBar.SetText("No Entries Found...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning) : objform.Close() : objrs = Nothing : Exit Sub
                Dim objDT As SAPbouiCOM.DataTable
                If objform.DataSources.DataTables.Count = 0 Then
                    objform.DataSources.DataTables.Add("DT_VIEW")
                End If

                objDT = objform.DataSources.DataTables.Item("DT_VIEW")
                objDT.ExecuteQuery(str_sql)
                objform.DataSources.DataTables.Item("DT_VIEW").ExecuteQuery(str_sql)

                Grid0.DataTable = objform.DataSources.DataTables.Item("DT_VIEW")
                Grid0.CollapseLevel = 1
                objform.Freeze(True)
                For i As Integer = 0 To Grid0.Columns.Count - 1
                    Grid0.Columns.Item(i).TitleObject.Sortable = True
                    Grid0.Columns.Item(i).Editable = False
                Next
                objform.Freeze(False)
                Grid0.Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                Dim col As SAPbouiCOM.EditTextColumn
                col = Grid0.Columns.Item(0)
                col.LinkedObjectType = LinkedID
                objform.Visible = True
                objform.Update()

            Catch ex As Exception
                objform.Freeze(False)
            End Try
        End Sub
        Private Sub Button0_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button0.ClickAfter
            Try
                objform.Close()
            Catch ex As Exception

            End Try

        End Sub
        Private WithEvents Button2 As SAPbouiCOM.Button
        Private WithEvents Button3 As SAPbouiCOM.Button

        Private Sub Button2_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button2.ClickAfter
            Try
                Grid0.Rows.ExpandAll()
                objaddon.objapplication.Menus.Item("1300").Activate()
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Button3_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button3.ClickAfter
            Try
                Grid0.Rows.CollapseAll()
                objaddon.objapplication.Menus.Item("1300").Activate()
            Catch ex As Exception

            End Try

        End Sub
    End Class
End Namespace
