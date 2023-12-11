Option Strict Off
Option Explicit On

Imports System.Drawing
Imports SAPbouiCOM.Framework

Namespace SubContractingPO
    <FormAttribute("SUBGEN", "SubContractingPO/GeneralSettings.b1f")>
    Friend Class GeneralSettings
        Inherits UserFormBase
        Private WithEvents objform As SAPbouiCOM.Form
        Private WithEvents pCFL As SAPbouiCOM.ISBOChooseFromListEventArg
        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.Button0 = CType(Me.GetItem("1").Specific, SAPbouiCOM.Button)
            Me.Button1 = CType(Me.GetItem("2").Specific, SAPbouiCOM.Button)
            Me.Folder1 = CType(Me.GetItem("fldrbom").Specific, SAPbouiCOM.Folder)
            Me.Folder2 = CType(Me.GetItem("fldrpo").Specific, SAPbouiCOM.Folder)
            Me.CheckBox0 = CType(Me.GetItem("DatePO").Specific, SAPbouiCOM.CheckBox)
            Me.CheckBox2 = CType(Me.GetItem("Reso").Specific, SAPbouiCOM.CheckBox)
            Me.CheckBox3 = CType(Me.GetItem("IBOM").Specific, SAPbouiCOM.CheckBox)
            Me.CheckBox4 = CType(Me.GetItem("POSel").Specific, SAPbouiCOM.CheckBox)
            Me.StaticText0 = CType(Me.GetItem("Item_11").Specific, SAPbouiCOM.StaticText)
            Me.EditText0 = CType(Me.GetItem("txtcode").Specific, SAPbouiCOM.EditText)
            Me.StaticText1 = CType(Me.GetItem("Item_13").Specific, SAPbouiCOM.StaticText)
            Me.EditText1 = CType(Me.GetItem("txtdoc").Specific, SAPbouiCOM.EditText)
            Me.CheckBox1 = CType(Me.GetItem("Cost").Specific, SAPbouiCOM.CheckBox)
            Me.CheckBox5 = CType(Me.GetItem("AutoProd").Specific, SAPbouiCOM.CheckBox)
            Me.CheckBox6 = CType(Me.GetItem("AutoRec").Specific, SAPbouiCOM.CheckBox)
            Me.CheckBox7 = CType(Me.GetItem("APLoad").Specific, SAPbouiCOM.CheckBox)
            Me.CheckBox8 = CType(Me.GetItem("SubScreen").Specific, SAPbouiCOM.CheckBox)
            Me.CheckBox9 = CType(Me.GetItem("Scrapcond").Specific, SAPbouiCOM.CheckBox)
            Me.EditText2 = CType(Me.GetItem("SGroup").Specific, SAPbouiCOM.EditText)
            Me.CheckBox10 = CType(Me.GetItem("CopyBOM").Specific, SAPbouiCOM.CheckBox)
            Me.StaticText2 = CType(Me.GetItem("GrpCode").Specific, SAPbouiCOM.StaticText)
            Me.CheckBox11 = CType(Me.GetItem("AutoItem").Specific, SAPbouiCOM.CheckBox)
            Me.CheckBox12 = CType(Me.GetItem("ToWhse").Specific, SAPbouiCOM.CheckBox)
            Me.EditText3 = CType(Me.GetItem("LCode").Specific, SAPbouiCOM.EditText)
            Me.StaticText3 = CType(Me.GetItem("LocCode").Specific, SAPbouiCOM.StaticText)
            Me.CheckBox13 = CType(Me.GetItem("Type").Specific, SAPbouiCOM.CheckBox)
            Me.Folder0 = CType(Me.GetItem("Tran").Specific, SAPbouiCOM.Folder)
            Me.CheckBox22 = CType(Me.GetItem("Price").Specific, SAPbouiCOM.CheckBox)
            Me.CheckBox23 = CType(Me.GetItem("WhseO").Specific, SAPbouiCOM.CheckBox)
            Me.EditText4 = CType(Me.GetItem("LCodeO").Specific, SAPbouiCOM.EditText)
            Me.StaticText4 = CType(Me.GetItem("LocInO").Specific, SAPbouiCOM.StaticText)
            Me.StaticText5 = CType(Me.GetItem("lGIcode").Specific, SAPbouiCOM.StaticText)
            Me.EditText5 = CType(Me.GetItem("tGIcode").Specific, SAPbouiCOM.EditText)
            Me.StaticText6 = CType(Me.GetItem("lGRcodeO").Specific, SAPbouiCOM.StaticText)
            Me.EditText6 = CType(Me.GetItem("tGRcodeO").Specific, SAPbouiCOM.EditText)
            Me.StaticText7 = CType(Me.GetItem("lGRcodeS").Specific, SAPbouiCOM.StaticText)
            Me.EditText7 = CType(Me.GetItem("tGRcodeS").Specific, SAPbouiCOM.EditText)
            Me.EditText9 = CType(Me.GetItem("tGIname").Specific, SAPbouiCOM.EditText)
            Me.EditText10 = CType(Me.GetItem("tGRnameO").Specific, SAPbouiCOM.EditText)
            Me.EditText11 = CType(Me.GetItem("tGRnameS").Specific, SAPbouiCOM.EditText)
            Me.LinkedButton0 = CType(Me.GetItem("lnkGI").Specific, SAPbouiCOM.LinkedButton)
            Me.LinkedButton1 = CType(Me.GetItem("lnkGRO").Specific, SAPbouiCOM.LinkedButton)
            Me.LinkedButton2 = CType(Me.GetItem("lnkGRS").Specific, SAPbouiCOM.LinkedButton)
            Me.CheckBox14 = CType(Me.GetItem("WhseITran").Specific, SAPbouiCOM.CheckBox)
            Me.EditText13 = CType(Me.GetItem("tInvWhse").Specific, SAPbouiCOM.EditText)
            Me.StaticText13 = CType(Me.GetItem("lInvWhse").Specific, SAPbouiCOM.StaticText)
            Me.CheckBox15 = CType(Me.GetItem("ChkTran").Specific, SAPbouiCOM.CheckBox)
            Me.EditText8 = CType(Me.GetItem("txtSUer").Specific, SAPbouiCOM.EditText)
            Me.StaticText8 = CType(Me.GetItem("lSUser").Specific, SAPbouiCOM.StaticText)
            Me.StaticText9 = CType(Me.GetItem("lblc1").Specific, SAPbouiCOM.StaticText)
            Me.StaticText10 = CType(Me.GetItem("lblc2").Specific, SAPbouiCOM.StaticText)
            Me.StaticText11 = CType(Me.GetItem("lblc3").Specific, SAPbouiCOM.StaticText)
            Me.CheckBox16 = CType(Me.GetItem("wprice").Specific, SAPbouiCOM.CheckBox)
            Me.StaticText12 = CType(Me.GetItem("lweight").Specific, SAPbouiCOM.StaticText)
            Me.CheckBox17 = CType(Me.GetItem("Field1").Specific, SAPbouiCOM.CheckBox)
            Me.CheckBox18 = CType(Me.GetItem("UDF0").Specific, SAPbouiCOM.CheckBox)
            Me.CheckBox19 = CType(Me.GetItem("UDF1").Specific, SAPbouiCOM.CheckBox)
            Me.OptionBtn0 = CType(Me.GetItem("RYes0").Specific, SAPbouiCOM.OptionBtn)
            Me.OptionBtn1 = CType(Me.GetItem("RNo0").Specific, SAPbouiCOM.OptionBtn)
            Me.StaticText15 = CType(Me.GetItem("Val0").Specific, SAPbouiCOM.StaticText)
            Me.OptionBtn2 = CType(Me.GetItem("RYes1").Specific, SAPbouiCOM.OptionBtn)
            Me.OptionBtn3 = CType(Me.GetItem("RNo1").Specific, SAPbouiCOM.OptionBtn)
            Me.StaticText16 = CType(Me.GetItem("Val1").Specific, SAPbouiCOM.StaticText)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()
            AddHandler ActivateAfter, AddressOf Me.Form_ActivateAfter

        End Sub
        Private WithEvents Button0 As SAPbouiCOM.Button
        Dim RecCount As String
        Private Sub OnCustomInitialize()
            Try
                objform = objaddon.objapplication.Forms.GetForm("SUBGEN", 0)
                objform = objaddon.objapplication.Forms.ActiveForm
                objform.Freeze(True)
                OptionBtn0.GroupWith("RNo0")
                OptionBtn2.GroupWith("RNo1")
                HeaderLabel_Color(StaticText9.Item, 11, Color.Purple.ToArgb, 15)
                HeaderLabel_Color(StaticText10.Item, 11, Color.Purple.ToArgb, 15)
                HeaderLabel_Color(StaticText11.Item, 11, Color.Purple.ToArgb, 15)
                If objaddon.HANA Then
                    objform.Items.Item("txtcode").Specific.String = objaddon.objglobalmethods.GetNextCode_Value("@MIPL_GEN")
                Else
                    objform.Items.Item("txtcode").Specific.String = objaddon.objglobalmethods.GetNextCode_Value("[@MIPL_GEN]")
                End If
                If objaddon.HANA Then
                    RecCount = objaddon.objglobalmethods.getSingleValue("select Count(*) from ""@MIPL_GEN"";")
                Else
                    RecCount = objaddon.objglobalmethods.getSingleValue("select Count(*) from [@MIPL_GEN]")
                End If
                StaticText1.Item.Visible = False
                objform.Items.Item("fldrbom").Click()
                FieldSize()
                EditText1.Item.Visible = False
                If RecCount = "1" Then
                    objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                    EditText0.Item.Enabled = True
                    EditText0.Value = "1"
                    'objform.ActiveItem = "txtdoc"
                    objform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    EditText0.Item.Enabled = False
                    Exit Sub
                End If
                objform.Freeze(False)
            Catch ex As Exception
            Finally
                objform.Freeze(False)
            End Try
        End Sub

        Private Sub HeaderLabel_Color(ByVal item As SAPbouiCOM.Item, ByVal fontsize As Integer, ByVal forecolor As Integer, ByVal height As Integer, Optional ByVal width As Integer = 0)
            item.TextStyle = FontStyle.Bold
            item.FontSize = fontsize
            'item.ForeColor = forecolor
            item.Height = height
            'If width <> 0 Then item.Width = width
        End Sub


        Private Sub FieldSize()
            Try
                CheckBox0.Item.Height = CheckBox0.Item.Height + 2
                CheckBox1.Item.Height = CheckBox1.Item.Height + 2
                CheckBox2.Item.Height = CheckBox2.Item.Height + 2
                CheckBox3.Item.Height = CheckBox3.Item.Height + 2
                CheckBox4.Item.Height = CheckBox4.Item.Height + 2
                CheckBox5.Item.Height = CheckBox5.Item.Height + 2
                CheckBox6.Item.Height = CheckBox6.Item.Height + 2
                CheckBox7.Item.Height = CheckBox7.Item.Height + 2
                CheckBox8.Item.Height = CheckBox8.Item.Height + 2
                CheckBox9.Item.Height = CheckBox9.Item.Height + 2
                CheckBox10.Item.Height = CheckBox10.Item.Height + 2
                CheckBox11.Item.Height = CheckBox11.Item.Height + 2
                CheckBox12.Item.Height = CheckBox12.Item.Height + 2
                CheckBox13.Item.Height = CheckBox13.Item.Height + 2
                CheckBox22.Item.Height = CheckBox22.Item.Height + 2
                CheckBox23.Item.Height = CheckBox23.Item.Height + 2
                CheckBox14.Item.Height = CheckBox14.Item.Height + 2
                CheckBox15.Item.Height = CheckBox15.Item.Height + 2
                CheckBox16.Item.Height = CheckBox16.Item.Height + 2
                CheckBox17.Item.Height = CheckBox17.Item.Height + 2
                CheckBox18.Item.Height = CheckBox18.Item.Height + 2
                CheckBox19.Item.Height = CheckBox19.Item.Height + 2

                OptionBtn0.Item.Height = OptionBtn0.Item.Height + 2
                OptionBtn1.Item.Height = OptionBtn1.Item.Height + 2
                OptionBtn2.Item.Height = OptionBtn2.Item.Height + 2
                OptionBtn3.Item.Height = OptionBtn3.Item.Height + 2

                CheckBox11.Item.Width = CheckBox11.Item.Width + 15
                CheckBox12.Item.Width = CheckBox12.Item.Width + 15
                CheckBox13.Item.Width = CheckBox13.Item.Width + 15
                CheckBox10.Item.Width = CheckBox10.Item.Width + 20
                CheckBox23.Item.Width = CheckBox23.Item.Width + 15

                OptionBtn0.Item.Width = OptionBtn0.Item.Width + 5
                OptionBtn1.Item.Width = OptionBtn1.Item.Width + 7
                OptionBtn2.Item.Width = OptionBtn2.Item.Width + 5
                OptionBtn3.Item.Width = OptionBtn3.Item.Width + 7

            Catch ex As Exception

            End Try
        End Sub
        Private WithEvents Button1 As SAPbouiCOM.Button
        Private WithEvents Folder1 As SAPbouiCOM.Folder
        Private WithEvents Folder2 As SAPbouiCOM.Folder
        Private WithEvents CheckBox0 As SAPbouiCOM.CheckBox
        Private WithEvents CheckBox2 As SAPbouiCOM.CheckBox
        Private WithEvents CheckBox3 As SAPbouiCOM.CheckBox
        Private WithEvents CheckBox4 As SAPbouiCOM.CheckBox
        Private WithEvents StaticText0 As SAPbouiCOM.StaticText
        Private WithEvents EditText0 As SAPbouiCOM.EditText
        Private WithEvents StaticText1 As SAPbouiCOM.StaticText
        Private WithEvents EditText1 As SAPbouiCOM.EditText

        Private Sub Button0_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button0.ClickAfter
            Try
                If objform.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                    If pVal.ActionSuccess Then
                        If objaddon.HANA Then
                            RecCount = objaddon.objglobalmethods.getSingleValue("select Count(*) +1 from ""@MIPL_GEN"";")
                        Else
                            RecCount = objaddon.objglobalmethods.getSingleValue("select Count(*) +1 from [@MIPL_GEN]")
                        End If
                        If RecCount <> "2" Then
                            objform.Close()
                        End If
                    End If
                End If

            Catch ex As Exception
            End Try

        End Sub

        Private WithEvents CheckBox1 As SAPbouiCOM.CheckBox
        Private WithEvents CheckBox5 As SAPbouiCOM.CheckBox
        Private WithEvents CheckBox6 As SAPbouiCOM.CheckBox
        Private WithEvents CheckBox7 As SAPbouiCOM.CheckBox
        Private WithEvents CheckBox8 As SAPbouiCOM.CheckBox
        Private WithEvents CheckBox9 As SAPbouiCOM.CheckBox
        Private WithEvents EditText2 As SAPbouiCOM.EditText

        Private Sub CheckBox9_PressedAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles CheckBox9.PressedAfter
            Try
                If CheckBox9.Checked = True Then
                    EditText2.Item.Enabled = True
                Else
                    EditText2.Item.Enabled = False
                End If
            Catch ex As Exception

            End Try

        End Sub
        Private WithEvents CheckBox10 As SAPbouiCOM.CheckBox
        Private WithEvents StaticText2 As SAPbouiCOM.StaticText
        Private WithEvents CheckBox11 As SAPbouiCOM.CheckBox
        Private WithEvents CheckBox12 As SAPbouiCOM.CheckBox
        Private WithEvents EditText3 As SAPbouiCOM.EditText
        Private WithEvents StaticText3 As SAPbouiCOM.StaticText

        Private Sub CheckBox12_PressedAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles CheckBox12.PressedAfter
            Try
                If CheckBox12.Checked = True Then
                    EditText3.Item.Enabled = True
                Else
                    EditText3.Item.Enabled = False
                End If
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Button0_ClickBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles Button0.ClickBefore
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then Exit Sub
                If pVal.InnerEvent = True Then BubbleEvent = False : Exit Sub
                If CheckBox9.Checked = True Then
                    If EditText2.Value = "" Then
                        objaddon.objapplication.StatusBar.SetText("Please update the ItemsGroup Code...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        BubbleEvent = False : Exit Sub
                    End If
                End If
                If CheckBox12.Checked = True Then
                    If EditText3.Value = "" Then
                        objaddon.objapplication.StatusBar.SetText("Please update the Location Code for Input Tab...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        BubbleEvent = False : Exit Sub
                    End If
                End If
                If CheckBox23.Checked = True Then
                    If EditText4.Value = "" Then
                        objaddon.objapplication.StatusBar.SetText("Please update the Location Code for Output Tab...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        BubbleEvent = False : Exit Sub
                    End If
                End If
                If CheckBox14.Checked = True Then
                    If EditText13.Value = "" Then
                        objaddon.objapplication.StatusBar.SetText("Please update the Warehouse Code for Posting Transfer...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        BubbleEvent = False : Exit Sub
                    End If
                End If
            Catch ex As Exception

            End Try

        End Sub
        Private WithEvents CheckBox13 As SAPbouiCOM.CheckBox
        Private WithEvents Folder0 As SAPbouiCOM.Folder
        Private WithEvents CheckBox22 As SAPbouiCOM.CheckBox
        Private WithEvents CheckBox23 As SAPbouiCOM.CheckBox
        Private WithEvents EditText4 As SAPbouiCOM.EditText
        Private WithEvents StaticText4 As SAPbouiCOM.StaticText

        Private Sub CheckBox23_PressedAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles CheckBox23.PressedAfter
            Try
                If pVal.ActionSuccess = False Then Exit Sub
                If CheckBox6.Checked = True Then
                    CheckBox23.Checked = False
                    objaddon.objapplication.StatusBar.SetText("Cannot get the warehouse in output tab since Receipt auto load selected...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Else
                    If CheckBox23.Checked = True Then
                        EditText4.Item.Enabled = True
                    Else
                        EditText4.Item.Enabled = False
                    End If
                End If
            Catch ex As Exception
            End Try

        End Sub

        Private Sub CheckBox11_PressedAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles CheckBox11.PressedAfter
            Try
                If pVal.ActionSuccess = False Then Exit Sub
                If CheckBox6.Checked = True Then
                    CheckBox11.Checked = False
                    objaddon.objapplication.StatusBar.SetText("Cannot auto populate the item since Receipt auto load selected...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Else
                    CheckBox11.Checked = True
                End If
            Catch ex As Exception
            End Try

        End Sub

        Private Sub CheckBox22_PressedAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles CheckBox22.PressedAfter
            Try
                If pVal.ActionSuccess = False Then Exit Sub
                If CheckBox6.Checked = True Then
                    CheckBox22.Checked = False
                    objaddon.objapplication.StatusBar.SetText("Cannot activate the vendor special price since Receipt auto load selected...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Else
                    CheckBox22.Checked = True
                End If
            Catch ex As Exception
            End Try

        End Sub

        Private Sub CheckBox6_PressedAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles CheckBox6.PressedAfter
            Try
                If pVal.ActionSuccess = False Then Exit Sub
                If CheckBox6.Checked = True Then
                    CheckBox22.Checked = False
                    CheckBox11.Checked = False
                    CheckBox23.Checked = False
                Else
                    'CheckBox22.Checked = True
                    'CheckBox11.Checked = True
                    'CheckBox23.Checked = True
                End If
            Catch ex As Exception

            End Try

        End Sub
        Private WithEvents StaticText5 As SAPbouiCOM.StaticText
        Private WithEvents EditText5 As SAPbouiCOM.EditText
        Private WithEvents StaticText6 As SAPbouiCOM.StaticText
        Private WithEvents EditText6 As SAPbouiCOM.EditText
        Private WithEvents StaticText7 As SAPbouiCOM.StaticText
        Private WithEvents EditText7 As SAPbouiCOM.EditText
        Private WithEvents EditText9 As SAPbouiCOM.EditText
        Private WithEvents EditText10 As SAPbouiCOM.EditText
        Private WithEvents EditText11 As SAPbouiCOM.EditText

        Private Sub CFLcondition(ByVal pVal As SAPbouiCOM.SBOItemEventArg, ByVal CFLID As String)
            If pVal.ActionSuccess = True Then Exit Sub
            Try
                Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item(CFLID)
                Dim oConds As SAPbouiCOM.Conditions
                Dim oCond As SAPbouiCOM.Condition
                Dim oEmptyConds As New SAPbouiCOM.Conditions
                oCFL.SetConditions(oEmptyConds)
                oConds = oCFL.GetConditions()

                oCond = oConds.Add()
                oCond.Alias = "Postable"
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCond.CondVal = "Y"
                oCFL.SetConditions(oConds)
            Catch ex As Exception
                SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Choose FromList Filter Pay Element Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End Try
        End Sub

        Private Sub ChooseFromList_AfterAction_AccountSelection(ByVal pVal As SAPbouiCOM.SBOItemEventArg, ByVal editext_acctcode As SAPbouiCOM.EditText, ByVal editext_acctname As SAPbouiCOM.EditText)
            Try
                If pVal.ActionSuccess = False Then Exit Sub
                pCFL = pVal
                If Not pCFL.SelectedObjects Is Nothing Then
                    Try
                        editext_acctcode.Value = pCFL.SelectedObjects.Columns.Item("AcctCode").Cells.Item(0).Value
                    Catch ex As Exception
                    End Try
                    Try
                        editext_acctname.Value = pCFL.SelectedObjects.Columns.Item("AcctName").Cells.Item(0).Value
                    Catch ex As Exception
                    End Try
                End If
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            Catch ex As Exception
            Finally
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        End Sub

        Private Sub EditText5_ChooseFromListBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles EditText5.ChooseFromListBefore
            Try 'Goods Issue    
                CFLcondition(pVal, "CFL_GI")
            Catch ex As Exception

            End Try

        End Sub

        Private Sub EditText5_ChooseFromListAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText5.ChooseFromListAfter
            Try 'Goods Issue    
                ChooseFromList_AfterAction_AccountSelection(pVal, EditText5, EditText9)
            Catch ex As Exception

            End Try

        End Sub
        Private WithEvents LinkedButton0 As SAPbouiCOM.LinkedButton
        Private WithEvents LinkedButton1 As SAPbouiCOM.LinkedButton
        Private WithEvents LinkedButton2 As SAPbouiCOM.LinkedButton

        Private Sub EditText6_ChooseFromListAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText6.ChooseFromListAfter
            Try 'Goods Receipt Output Tab    
                ChooseFromList_AfterAction_AccountSelection(pVal, EditText6, EditText10)
            Catch ex As Exception

            End Try

        End Sub

        Private Sub EditText6_ChooseFromListBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles EditText6.ChooseFromListBefore
            Try 'Goods Receipt Output Tab    
                CFLcondition(pVal, "CFL_GRO")
            Catch ex As Exception

            End Try


        End Sub

        Private Sub EditText7_ChooseFromListBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles EditText7.ChooseFromListBefore
            Try 'Goods Receipt Scrap Tab    
                CFLcondition(pVal, "CFL_GRS")
            Catch ex As Exception

            End Try

        End Sub

        Private Sub EditText7_ChooseFromListAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText7.ChooseFromListAfter
            Try 'Goods Receipt Scrap Tab    
                ChooseFromList_AfterAction_AccountSelection(pVal, EditText7, EditText11)
            Catch ex As Exception

            End Try

        End Sub

        Private WithEvents CheckBox14 As SAPbouiCOM.CheckBox
        Private WithEvents EditText13 As SAPbouiCOM.EditText
        Private WithEvents StaticText13 As SAPbouiCOM.StaticText

        Private Sub CheckBox14_PressedAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles CheckBox14.PressedAfter
            Try
                If CheckBox14.Checked = True Then
                    EditText13.Item.Enabled = True
                Else
                    EditText13.Item.Enabled = False
                End If
            Catch ex As Exception

            End Try

        End Sub
        Private WithEvents CheckBox15 As SAPbouiCOM.CheckBox


        Private Sub EditText2_ChooseFromListAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText2.ChooseFromListAfter
            'GrpCode
            Try
                If pVal.ActionSuccess = False Then Exit Sub
                pCFL = pVal
                If Not pCFL.SelectedObjects Is Nothing Then
                    Try
                        EditText2.Value = pCFL.SelectedObjects.Columns.Item("ItmsGrpCod").Cells.Item(0).Value
                    Catch ex As Exception
                    End Try
                End If
            Catch ex As Exception
            Finally
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        End Sub
        Private WithEvents EditText8 As SAPbouiCOM.EditText
        Private WithEvents StaticText8 As SAPbouiCOM.StaticText

        Private Sub EditText8_ChooseFromListAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText8.ChooseFromListAfter
            Try
                If pVal.ActionSuccess = False Then Exit Sub
                pCFL = pVal
                If Not pCFL.SelectedObjects Is Nothing Then
                    Try
                        EditText8.Value = pCFL.SelectedObjects.Columns.Item("USER_CODE").Cells.Item(0).Value
                    Catch ex As Exception
                    End Try
                End If
            Catch ex As Exception
            Finally
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try

        End Sub

        Private Sub Form_ActivateAfter(pVal As SAPbouiCOM.SBOItemEventArg)
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            Catch ex As Exception

            End Try

        End Sub

        Private Sub EditText13_ChooseFromListAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText13.ChooseFromListAfter
            Try
                If pVal.ActionSuccess = False Then Exit Sub
                pCFL = pVal
                If Not pCFL.SelectedObjects Is Nothing Then
                    Try
                        EditText13.Value = pCFL.SelectedObjects.Columns.Item("WhsCode").Cells.Item(0).Value
                    Catch ex As Exception
                    End Try
                End If
            Catch ex As Exception
            Finally
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try

        End Sub

        Private Sub EditText3_KeyDownAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText3.KeyDownAfter
            Try
                If pVal.CharPressed = 13 Or pVal.CharPressed = 9 Then
                    If objaddon.HANA Then
                        ProductQuery = "Select Distinct 'N' ""Select"",D.""Code"", D.""Location"" From OLCT D "
                    Else
                        ProductQuery = "Select Distinct 'N' Select,D.Code, D.Location From OLCT D "
                    End If
                    FrmMultiSel = objaddon.objapplication.Forms.ActiveForm
                    If Not objaddon.FormExist("MULSEL") Then
                        FieldName = "LCode"
                        Dim Multiselect As New FrmMultiSelect
                        Multiselect.Show()
                    End If
                End If
            Catch ex As Exception

            End Try

        End Sub

        Private Sub EditText4_KeyDownAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText4.KeyDownAfter
            Try
                If pVal.CharPressed = 13 Or pVal.CharPressed = 9 Then
                    If objaddon.HANA Then
                        ProductQuery = "Select Distinct 'N' ""Select"",D.""Code"", D.""Location"" From OLCT D "
                    Else
                        ProductQuery = "Select Distinct 'N' Select,D.Code, D.Location From OLCT D "
                    End If

                    FrmMultiSel = objaddon.objapplication.Forms.ActiveForm
                    If Not objaddon.FormExist("MULSEL") Then
                        FieldName = "LCodeO"
                        Dim Multiselect As New FrmMultiSelect
                        Multiselect.Show()
                    End If
                End If
            Catch ex As Exception

            End Try
        End Sub

        Private WithEvents StaticText9 As SAPbouiCOM.StaticText
        Private WithEvents StaticText10 As SAPbouiCOM.StaticText
        Private WithEvents StaticText11 As SAPbouiCOM.StaticText
        Private WithEvents CheckBox16 As SAPbouiCOM.CheckBox
        Private WithEvents StaticText12 As SAPbouiCOM.StaticText
        Private WithEvents CheckBox17 As SAPbouiCOM.CheckBox
        Private WithEvents CheckBox18 As SAPbouiCOM.CheckBox
        Private WithEvents CheckBox19 As SAPbouiCOM.CheckBox
        Private WithEvents OptionBtn0 As SAPbouiCOM.OptionBtn
        Private WithEvents OptionBtn1 As SAPbouiCOM.OptionBtn
        Private WithEvents StaticText15 As SAPbouiCOM.StaticText
        Private WithEvents OptionBtn2 As SAPbouiCOM.OptionBtn
        Private WithEvents OptionBtn3 As SAPbouiCOM.OptionBtn
        Private WithEvents StaticText16 As SAPbouiCOM.StaticText

        Private Sub OptionBtn0_PressedAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles OptionBtn0.PressedAfter
            Try
                If Not CheckBox18.Checked = True Then
                    OptionBtn1.Selected = True
                End If
            Catch ex As Exception

            End Try

        End Sub

        Private Sub OptionBtn2_PressedAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles OptionBtn2.PressedAfter
            Try
                If Not CheckBox19.Checked = True Then
                    OptionBtn3.Selected = True
                End If
            Catch ex As Exception

            End Try


        End Sub
    End Class
End Namespace
