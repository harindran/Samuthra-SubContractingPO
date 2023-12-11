Namespace SubContractingPO

    Public Class clsRightClickEvent
        Dim objform As SAPbouiCOM.Form
        Dim objglobalmethods As New clsGlobalMethods
        Dim ocombo As SAPbouiCOM.ComboBox
        Dim objmatrix As SAPbouiCOM.Matrix
        Dim strsql As String
        Dim objrs As SAPbobsCOM.Recordset

        Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
            Try
                Select Case objaddon.objapplication.Forms.ActiveForm.TypeEx
                    Case "SUBCTPO"
                        SubContractingPO_RightClickEvent(eventInfo, BubbleEvent)
                    Case "SUBBOM"
                        SubContractingBOM_RightClickEvent(eventInfo, BubbleEvent)
                    Case "65211"
                        ProductionOrder_RightClickEvent(eventInfo, BubbleEvent)
                    Case "SUBGEN"
                        GeneralSettings_RightClickEvent(eventInfo, BubbleEvent)
                End Select
            Catch ex As Exception
            End Try
        End Sub

        Private Sub RightClickMenu_Add(ByVal MainMenu As String, ByVal NewMenuID As String, ByVal NewMenuName As String, ByVal position As Integer)
            Dim omenus As SAPbouiCOM.Menus
            Dim omenuitem As SAPbouiCOM.MenuItem
            Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
            oCreationPackage = objaddon.objapplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            omenuitem = objaddon.objapplication.Menus.Item(MainMenu) 'Data'
            If Not omenuitem.SubMenus.Exists(NewMenuID) Then
                oCreationPackage.UniqueID = NewMenuID
                oCreationPackage.String = NewMenuName
                oCreationPackage.Position = position
                oCreationPackage.Enabled = True
                omenus = omenuitem.SubMenus
                omenus.AddEx(oCreationPackage)
            End If
        End Sub

        Private Sub RightClickMenu_Delete(ByVal MainMenu As String, ByVal NewMenuID As String)
            Dim omenuitem As SAPbouiCOM.MenuItem
            omenuitem = objaddon.objapplication.Menus.Item(MainMenu) 'Data'
            If omenuitem.SubMenus.Exists(NewMenuID) Then
                objaddon.objapplication.Menus.RemoveEx(NewMenuID)
            End If
        End Sub

        Private Sub SubContractingPO_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
            Try
                Dim objform As SAPbouiCOM.Form
                objform = objaddon.objapplication.Forms.ActiveForm
                Dim Matrix2, Matrix3, objMatrix As SAPbouiCOM.Matrix
                Dim FolderOutput, FolderScrap As SAPbouiCOM.Folder
                Matrix2 = objform.Items.Item("mtxscrap").Specific
                FolderOutput = objform.Items.Item("flroutput").Specific
                Matrix3 = objform.Items.Item("mtxoutput").Specific
                FolderScrap = objform.Items.Item("flrscrap").Specific
                If eventInfo.BeforeAction Then
                    objform.EnableMenu("1283", False)
                    If eventInfo.ItemUID <> "" Then
                        Try
                            objMatrix = objform.Items.Item(eventInfo.ItemUID).Specific
                            If objMatrix.Item.Type = SAPbouiCOM.BoFormItemTypes.it_MATRIX Then
                                If objMatrix.Columns.Item(eventInfo.ColUID).Cells.Item(eventInfo.Row).Specific.String <> "" Then
                                    objform.EnableMenu("772", True)  'Copy
                                Else
                                    objform.EnableMenu("772", False)
                                End If
                            End If
                        Catch ex As Exception
                            If objform.Items.Item(eventInfo.ItemUID).Specific.String <> "" Then
                                objform.EnableMenu("772", True)  'Copy
                            Else
                                objform.EnableMenu("772", False)
                            End If
                        End Try
                    Else
                        objform.EnableMenu("772", False)
                    End If

                    objform.EnableMenu("784", True)  'Copy Table
                    If eventInfo.ItemUID = "MtxCosting" Then
                        objform.EnableMenu("1292", True) 'Add Row Menu
                        objform.EnableMenu("1293", True) 'Remove Row Menu
                        objform.EnableMenu("784", True)  'Copy Table
                    ElseIf eventInfo.ItemUID = "MtxinputN" Then
                        If ItemBOM = "Y" Then
                            objform.EnableMenu("1293", False) 'Remove Row Menu
                        Else
                            objform.EnableMenu("1293", True) 'Remove Row Menu
                        End If
                        'If objform.Items.Item("TxtInvTr").Specific.String <> "" Then
                        '    objform.EnableMenu("1293", False) 'Remove Row Menu
                        'Else
                        '    objform.EnableMenu("1293", True) 'Remove Row Menu
                        'End If
                    ElseIf eventInfo.ItemUID = "mtxoutput" Or eventInfo.ItemUID = "mtxscrap" Then
                        Matrix3.Columns.Item("Quant").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                        Matrix2.Columns.Item("Quant").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                        'objform.EnableMenu("784", True)  'Copy Table
                        If FolderOutput.Selected = True Then
                            If Not ReceiptAutoLoad = "Y" Then
                                If objaddon.objglobalmethods.AutoAssign_SubItem(FolderOutput, Matrix3) Then
                                    Dim Qty As Double = 0
                                    For i As Integer = 1 To Matrix3.VisualRowCount
                                        If Matrix3.Columns.Item("Code").Cells.Item(i).Specific.string <> "" And Matrix3.Columns.Item("Code").Cells.Item(i).Specific.string = objform.Items.Item("txtsitem").Specific.String Then
                                            Qty += CDbl(Matrix3.Columns.Item("Quant").Cells.Item(i).Specific.string)
                                        End If
                                    Next
                                    If Qty < CDbl(objform.Items.Item("SQty").Specific.String) Then
                                        objform.EnableMenu("1292", True) 'Add Row Menu
                                    End If
                                End If
                            End If
                            If Matrix3.Columns.Item("Code").Cells.Item(eventInfo.Row).Specific.String <> "" And Matrix3.Columns.Item("GINo").Cells.Item(eventInfo.Row).Specific.String <> "" Or Matrix3.Columns.Item("GRNo").Cells.Item(eventInfo.Row).Specific.String <> "" Then
                                objform.EnableMenu("1293", False) 'Remove Row Menu
                            Else
                                objform.EnableMenu("1293", True) 'Remove Row Menu
                            End If
                        ElseIf FolderScrap.Selected = True Then
                            'If objaddon.objglobalmethods.AutoAssign_SubItem(FolderScrap, Matrix2) Then
                            '    objform.EnableMenu("1292", True) 'Add Row Menu
                            'End If
                            If Matrix2.Columns.Item("Code").Cells.Item(eventInfo.Row).Specific.String <> "" Then
                                objform.EnableMenu("1292", True) 'Add Row Menu
                            End If
                            objform.EnableMenu("1293", True) 'Remove Row Menu
                            'If Matrix2.Columns.Item("Code").Cells.Item(eventInfo.Row).Specific.String <> "" And Matrix2.Columns.Item("GRNo").Cells.Item(eventInfo.Row).Specific.String <> "" Or Matrix2.Columns.Item("InvNum").Cells.Item(eventInfo.Row).Specific.String <> "" Then
                            'If Matrix2.Columns.Item("Code").Cells.Item(eventInfo.Row).Specific.String <> "" Then
                            '    objform.EnableMenu("1293", False) 'Remove Row Menu
                            'Else
                            '    objform.EnableMenu("1293", True) 'Remove Row Menu
                            'End If
                        End If
                    ElseIf eventInfo.ItemUID = "mtxreldoc" Then
                        objform.EnableMenu("1292", True) 'Add Row Menu
                        objform.EnableMenu("1293", True) 'Remove Row Menu
                        'objform.EnableMenu("784", True)  'Copy Table
                    End If

                        If objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                            objform.EnableMenu("1287", True)  'Duplicate
                        Else
                            objform.EnableMenu("1287", False)
                        End If
                Else
                        objform.EnableMenu("1292", False) 'Add Row Menu
                        objform.EnableMenu("1293", False) 'Remove Row Menu
                        objform.EnableMenu("1283", False)
                        objform.EnableMenu("784", False)
                    End If
            Catch ex As Exception
            End Try
        End Sub

        Private Sub SubContractingBOM_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
            Try
                Dim objform As SAPbouiCOM.Form
                objform = objaddon.objapplication.Forms.ActiveForm
                If eventInfo.BeforeAction Then
                    objform.EnableMenu("1283", True)
                    If eventInfo.ItemUID <> "" Then
                        Try
                            objmatrix = objform.Items.Item(eventInfo.ItemUID).Specific
                            If objmatrix.Item.Type = SAPbouiCOM.BoFormItemTypes.it_MATRIX Then
                                If objmatrix.Columns.Item(eventInfo.ColUID).Cells.Item(eventInfo.Row).Specific.String <> "" Then
                                    objform.EnableMenu("772", True)  'Copy
                                Else
                                    objform.EnableMenu("772", False)
                                End If
                            End If
                        Catch ex As Exception
                            If objform.Items.Item(eventInfo.ItemUID).Specific.String <> "" Then
                                objform.EnableMenu("772", True)  'Copy
                            Else
                                objform.EnableMenu("772", False)
                            End If
                        End Try
                    Else
                        objform.EnableMenu("772", False)
                    End If

                    If eventInfo.ItemUID = "mtxBOM" Then
                        objform.EnableMenu("1292", True) 'Add Row Menu
                        objform.EnableMenu("1293", True) 'Remove Row Menu
                        objform.EnableMenu("784", True)  'Copy Table
                    Else
                        objform.EnableMenu("1292", False) 'Add Row Menu
                        objform.EnableMenu("1293", False) 'Remove Row Menu
                    End If
                Else
                    objform.EnableMenu("1283", False)
                    objform.EnableMenu("784", False)
                End If
            Catch ex As Exception
            End Try
        End Sub

        Private Sub GeneralSettings_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
            Try
                Dim objform As SAPbouiCOM.Form
                objform = objaddon.objapplication.Forms.ActiveForm
                If eventInfo.BeforeAction Then
                    objform.EnableMenu("1283", False)
                    objform.EnableMenu("784", False)
                    'Else
                    '    objform.EnableMenu("1283", False)
                    '    objform.EnableMenu("784", False)
                End If
            Catch ex As Exception
            End Try
        End Sub

        Private Sub ProductionOrder_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
            Try
                Dim objform As SAPbouiCOM.Form
                Dim StrQuery As String
                Dim oMenuItem As SAPbouiCOM.MenuItem
                objform = objaddon.objapplication.Forms.ActiveForm
                If objaddon.HANA Then
                    SubPOScreen = objaddon.objglobalmethods.getSingleValue("select ""U_SubScreen"" from ""@MIPL_GEN"" where ifnull(""U_SubScreen"",'')='Y'")
                Else
                    SubPOScreen = objaddon.objglobalmethods.getSingleValue("select U_SubScreen from [@MIPL_GEN] where isnull(U_SubScreen,'')='Y'")
                End If

                If SubPOScreen <> "Y" Then
                    If objaddon.HANA Then
                        StrQuery = objaddon.objglobalmethods.getSingleValue("select ""U_SubConBOM"" from OITT where ""Code""='" & objform.Items.Item("6").Specific.String & "'")
                    Else
                        StrQuery = objaddon.objglobalmethods.getSingleValue("select U_SubConBOM from OITT where Code='" & objform.Items.Item("6").Specific.String & "'")
                    End If
                    If eventInfo.BeforeAction Then
                        If StrQuery = "Y" Then
                            oMenuItem = objaddon.objapplication.Menus.Item("5922")
                            oMenuItem.Enabled = False
                            oMenuItem = objaddon.objapplication.Menus.Item("5923")
                            oMenuItem.Enabled = False
                        Else
                            oMenuItem = objaddon.objapplication.Menus.Item("5922")
                            oMenuItem.Enabled = True
                            oMenuItem = objaddon.objapplication.Menus.Item("5923")
                            oMenuItem.Enabled = True
                        End If
                    Else
                    End If
                Else

                End If

            Catch ex As Exception

            End Try
        End Sub
    End Class

End Namespace
