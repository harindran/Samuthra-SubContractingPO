Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework

Namespace SubContractingPO
    <FormAttribute("MULSEL", "SubContractingPO/FrmMultiSelect.b1f")>
    Friend Class FrmMultiSelect
        Inherits UserFormBase
        Private WithEvents objform As SAPbouiCOM.Form
        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.Button0 = CType(Me.GetItem("101").Specific, SAPbouiCOM.Button)
            Me.Button1 = CType(Me.GetItem("2").Specific, SAPbouiCOM.Button)
            Me.Grid0 = CType(Me.GetItem("Grdata").Specific, SAPbouiCOM.Grid)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()

        End Sub

        Private WithEvents Button0 As SAPbouiCOM.Button

        Private Sub OnCustomInitialize()
            Try
                objform = objaddon.objapplication.Forms.GetForm("MULSEL", 0)
                LoadGrid(ProductQuery)

                objform.Update()
                objform.Refresh()
            Catch ex As Exception

            End Try

        End Sub

        Private WithEvents Button1 As SAPbouiCOM.Button
        Private WithEvents Grid0 As SAPbouiCOM.Grid

        Private Sub LoadGrid(ByVal query As String)
            Try

                Grid0.DataTable.ExecuteQuery(query)
                Grid0.Columns.Item("Select").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
                Grid0.Columns.Item("Code").Editable = False
                Grid0.Columns.Item("Location").Editable = False
                'Grid0.Columns.Item("#").Editable = False
                ProductQuery = ""
            Catch ex As Exception

            End Try
        End Sub

        Private Sub Button0_ClickAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button0.ClickAfter
            Try
                Dim code As String = ""
                For i As Integer = 0 To Grid0.Rows.Count - 1
                    If Grid0.DataTable.GetValue("Select", i).ToString = "Y" Then
                        code = code + (Grid0.DataTable.GetValue(1, i).ToString + ",")
                    End If
                Next
                code = code.Remove(code.Length - 1)
                If code <> "" Then
                    FrmMultiSel.Items.Item(FieldName).Specific.string = ""
                    FrmMultiSel.Items.Item(FieldName).Specific.string = code
                End If
                FrmMultiSel = Nothing
                FieldName = Nothing
                objform.Close()
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Grid0_DoubleClickAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Grid0.DoubleClickAfter
            Try

                Select Case pVal.ColUID
                    Case "Select"
                        If pVal.Row = -1 Then
                            'objform.Freeze(True)
                            If Grid0.DataTable.GetValue("Select", 0).ToString = "Y" Then
                                For i As Integer = 0 To Grid0.Rows.Count - 1
                                    Grid0.Columns.Item("Select").Check(i, False)
                                Next
                            Else
                                For i As Integer = 0 To Grid0.Rows.Count - 1
                                    Grid0.Columns.Item("Select").Check(i, True)
                                Next
                            End If

                            'objform.Freeze(False)
                        End If
                    Case "Code"
                        If pVal.Row = -1 Then
                            Grid0.Columns.Item("Code").TitleObject.Sortable = True
                        End If
                    Case "Name"
                        If pVal.Row = -1 Then
                            Grid0.Columns.Item("Name").TitleObject.Sortable = True
                            'Grid0.Columns.Item("Name").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
                        End If
                End Select
            Catch ex As Exception

            End Try

        End Sub
    End Class
End Namespace
