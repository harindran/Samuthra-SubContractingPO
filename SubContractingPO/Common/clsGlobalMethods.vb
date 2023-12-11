Imports System.Text.RegularExpressions
Imports System.Drawing

Namespace SubContractingPO

    Public Class clsGlobalMethods
        Dim strsql As String
        Dim objrs As SAPbobsCOM.Recordset

        Public Function GetDocNum(ByVal sUDOName As String, ByVal Series As Integer) As String
            Dim StrSQL As String
            Dim objRS As SAPbobsCOM.Recordset
            objRS = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'If objAddOn.HANA Then
            If Series = 0 Then
                StrSQL = " select  ""NextNumber""  from NNM1 where ""ObjectCode""='" & sUDOName & "'"
            Else
                StrSQL = " select  ""NextNumber""  from NNM1 where ""ObjectCode""='" & sUDOName & "' and ""Series"" = " & Series
            End If

            'Else
            'StrSQL = "select Autokey from onnm where objectcode='" & sUDOName & "'"
            'End If
            objRS.DoQuery(StrSQL)
            objRS.MoveFirst()
            If Not objRS.EoF Then
                Return Convert.ToInt32(objRS.Fields.Item(0).Value.ToString())
            Else
                GetDocNum = "1"
            End If
        End Function
        Public Function GetNextCode_Value(ByVal Tablename As String)
            Try
                If Tablename.ToString = "" Then Return ""
                If objaddon.HANA Then
                    strsql = "select IFNULL(Max(CAST(""Code"" As integer)),0)+1 from """ & Tablename.ToString & """"
                Else
                    strsql = "select ISNULL(Max(CAST(Code As integer)),0)+1 from " & Tablename.ToString & ""
                End If

                objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objrs.DoQuery(strsql)
                If objrs.RecordCount > 0 Then Return objrs.Fields.Item(0).Value.ToString Else Return ""
            Catch ex As Exception
                objaddon.objapplication.SetStatusBarMessage("Error while getting next code numbe" & ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                Return ""
            End Try
        End Function

        Public Function GetNextDocNum_Value(ByVal Tablename As String)
            Try
                If Tablename.ToString = "" Then Return ""
                strsql = "select IFNULL(Max(CAST(""DocNum"" As integer)),0)+1 from """ & Tablename.ToString & """"
                objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objrs.DoQuery(strsql)
                If objrs.RecordCount > 0 Then Return objrs.Fields.Item(0).Value.ToString Else Return ""
            Catch ex As Exception
                objaddon.objapplication.SetStatusBarMessage("Error while getting next code numbe" & ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                Return ""
            End Try
        End Function
        Public Function GetNextDocEntry_Value(ByVal Tablename As String)
            Try
                If Tablename.ToString = "" Then Return ""
                strsql = "select IFNULL(Max(CAST(""DocEntry"" As integer)),0)+1 from """ & Tablename.ToString & """"
                objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objrs.DoQuery(strsql)
                If objrs.RecordCount > 0 Then Return objrs.Fields.Item(0).Value.ToString Else Return ""
            Catch ex As Exception
                objaddon.objapplication.SetStatusBarMessage("Error while getting next code numbe" & ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                Return ""
            End Try
        End Function

        Public Function GetDuration_BetWeenTime(ByVal strFrom As String, ByVal strTo As String)
            Dim Fromtime, Totime As DateTime
            Dim Duration As TimeSpan
            strFrom = Convert_String_TimeHHMM(strFrom) : strTo = Convert_String_TimeHHMM(strTo)
            Totime = New DateTime(2000, 1, 1, Left(strTo, 2), Right(strTo, 2), 0)
            Fromtime = New DateTime(2000, 1, 1, Left(strFrom, 2), Right(strFrom, 2), 0)
            If Totime < Fromtime Then Totime = New DateTime(2000, 1, 2, Left(strTo, 2), Right(strTo, 2), 0)
            Duration = Totime - Fromtime
            Return Duration.Hours.ToString + "." + Left((Duration.Minutes.ToString + "00"), 2).ToString
        End Function
        Public Function GetHours(ByVal FromHrs As String, ByVal ToHrs As String)
            Dim StartTime = New DateTime(2001, 1, 1, FromHrs, 0, 0)
            Dim EndTime = New DateTime(2001, 1, 1, ToHrs, 0, 0)
            Dim duration = EndTime - StartTime
            Dim durationhr = duration.TotalHours '+ "." + duration.TotalMinutes
            Return durationhr
        End Function
        Public Function Validation_From_To_Time(ByVal strFrom As String, ByVal strTo As String)
            Dim Fromtime, Totime As DateTime
            Dim Duration As TimeSpan
            strFrom = Convert_String_TimeHHMM(strFrom) : strTo = Convert_String_TimeHHMM(strTo)
            Totime = New DateTime(2000, 1, 1, Left(strTo, 2), Right(strTo, 2), 0)
            Fromtime = New DateTime(2000, 1, 1, Left(strFrom, 2), Right(strFrom, 2), 0)
            If Totime < Fromtime Then Totime = New DateTime(2000, 1, 2, Left(strTo, 2), Right(strTo, 2), 0)
            Duration = Totime - Fromtime
            If Duration.Hours < 0 Or Duration.Minutes < 0 Then Return False
            Return True
        End Function

        Public Function Convert_String_TimeHHMM(ByVal str As String)
            Return Right("0000" + Regex.Replace(str, "[^\d]", ""), 4)
        End Function

        Public Sub LoadCombo(ByVal objcombo As SAPbouiCOM.ComboBox, Optional ByVal strquery As String = "", Optional ByVal rs As SAPbobsCOM.Recordset = Nothing)
            objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            If strquery.ToString = "" And rs Is Nothing Then Exit Sub
            If strquery.ToString <> "" Then objrs.DoQuery(strquery) Else objrs = rs
            If objrs.RecordCount = 0 Then Exit Sub : If objrs.Fields.Count < 2 Then Exit Sub

            If objcombo.ValidValues.Count > 0 Then
                For i As Integer = objcombo.ValidValues.Count - 1 To 0 Step -1
                    objcombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                Next
            End If

            objrs.MoveFirst()
            For i As Integer = 0 To objrs.RecordCount - 1
                objcombo.ValidValues.Add(objrs.Fields.Item(0).Value.ToString, objrs.Fields.Item(1).Value.ToString)
                objrs.MoveNext()
            Next
        End Sub

        Public Sub LoadCombo_Series(ByVal objform As SAPbouiCOM.Form, ByVal comboname As String, ByVal objectid As String, ByVal docdate As Date)
            Try
                Dim objcombo As SAPbouiCOM.ComboBox
                If objectid.ToString = "" Then Exit Sub
                objcombo = objform.Items.Item(comboname).Specific
                'objcombo.ValidValues.LoadSeries(objectid, SAPbouiCOM.BoSeriesMode.sf_Add)

                objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'strsql = " Select Series,Seriesname from nnm1 where objectcode='" & objectid.ToString & "' and Indicator in (select Distinct Indicator  from OFPR where PeriodStat <>'Y') "
                strsql = " CALL ""MIPL_GetDefaultSeries"" ('" & objectid.ToString & "','" & objaddon.objcompany.UserName & "','" & docdate.ToString("yyyyMMdd") & "')"
                objrs.DoQuery(strsql)

                If objrs.RecordCount = 0 Then Exit Sub : If objrs.Fields.Count < 2 Then Exit Sub

                If objcombo.ValidValues.Count > 0 Then
                    For i As Integer = objcombo.ValidValues.Count - 1 To 0 Step -1 : objcombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index) : Next
                End If

                objrs.MoveFirst()
                For i As Integer = 0 To objrs.RecordCount - 1
                    objcombo.ValidValues.Add(objrs.Fields.Item(0).Value.ToString, objrs.Fields.Item(1).Value.ToString)
                    objrs.MoveNext()
                Next

                objrs.MoveFirst()
                objcombo.Select(objrs.Fields.Item("dflt").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)

            Catch ex As Exception

            End Try
        End Sub

        Public Sub LoadCombo_SingleSeries_AfterFind(ByVal objform As SAPbouiCOM.Form, ByVal comboname As String, ByVal objectid As String, ByVal Seriesid As String)
            Try
                If objectid.ToString = "" Or Seriesid = "" Or comboname = "" Or objform Is Nothing Then Exit Sub

                Dim objcombo As SAPbouiCOM.ComboBox
                objcombo = objform.Items.Item(comboname).Specific
                'objcombo.ValidValues.LoadSeries(objectid, SAPbouiCOM.BoSeriesMode.sf_Add)

                objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                strsql = " Select ""Series"",""SeriesName"" from nnm1 where ""ObjectCode""='" & objectid.ToString & "' and ""Series""='" & Seriesid.ToString & "'"
                objrs.DoQuery(strsql)

                If objrs.RecordCount = 0 Then Exit Sub : If objrs.Fields.Count < 2 Then Exit Sub

                'If objcombo.ValidValues.Count > 0 Then
                '    For i As Integer = objcombo.ValidValues.Count - 1 To 0 Step -1 : objcombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index) : Next
                'End If

                objcombo.ValidValues.Add(objrs.Fields.Item(0).Value.ToString, objrs.Fields.Item(1).Value.ToString)

                objcombo.Select(Seriesid, SAPbouiCOM.BoSearchKey.psk_ByValue)

            Catch ex As Exception

            End Try
        End Sub

        Public Function default_series(ByVal objectid As String, ByVal docdate As Date)
            Try
                objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                strsql = " CALL ""MIPL_GetDefaultSeries"" ('" & objectid.ToString & "','" & objaddon.objcompany.UserName & "','" & docdate & "')"
                objrs.DoQuery(strsql)

                If objrs.RecordCount = 0 Then Return ""
                Return objrs.Fields.Item(0).Value
            Catch ex As Exception
                Return ""
            End Try
        End Function

        Public Sub Matrix_Addrow(ByVal omatrix As SAPbouiCOM.Matrix, Optional ByVal colname As String = "", Optional ByVal rowno_name As String = "", Optional ByVal Error_Needed As Boolean = False)
            Try
                Dim addrow As Boolean = False

                If omatrix.VisualRowCount = 0 Then addrow = True : GoTo addrow
                If colname = "" Then addrow = True : GoTo addrow
                If omatrix.Columns.Item(colname).Cells.Item(omatrix.VisualRowCount).Specific.string <> "" Then addrow = True : GoTo addrow

addrow:
                If addrow = True Then
                    omatrix.AddRow(1)
                    omatrix.ClearRowData(omatrix.VisualRowCount)
                    If rowno_name <> "" Then omatrix.Columns.Item("#").Cells.Item(omatrix.VisualRowCount).Specific.string = omatrix.VisualRowCount
                Else
                    If Error_Needed = True Then objaddon.objapplication.SetStatusBarMessage("Already Empty Row Available", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                End If
            Catch ex As Exception

            End Try
        End Sub

        Public Sub HeaderLabel_Color(ByRef item As SAPbouiCOM.Item, ByVal fontsize As Integer, ByVal forecolor As Integer, ByVal height As Integer, Optional ByVal width As Integer = 0)
            item.TextStyle = FontStyle.Bold
            item.FontSize = fontsize
            item.ForeColor = forecolor
            item.Height = height
            'If width <> 0 Then item.Width = width
        End Sub

        Public Sub RightClickMenu_Delete(ByVal MainMenu As String, ByVal NewMenuID As String)
            Try
                Dim omenuitem As SAPbouiCOM.MenuItem
                omenuitem = objaddon.objapplication.Menus.Item(MainMenu) 'Data'
                If omenuitem.SubMenus.Exists(NewMenuID) Then
                    objaddon.objapplication.Menus.RemoveEx(NewMenuID)
                End If
            Catch ex As Exception

            End Try
        End Sub

        Public Sub SetAutomanagedattribute_Editable(ByVal oform As SAPbouiCOM.Form, ByVal fieldname As String, ByVal add As Boolean, ByVal find As Boolean, ByVal update As Boolean)

            If add = True Then
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            Else
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            End If

            If find = True Then
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            Else
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            End If

            If update Then
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            Else
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            End If
        End Sub

        Public Sub SetAutomanagedattribute_Visible(ByVal oform As SAPbouiCOM.Form, ByVal fieldname As String, ByVal add As Boolean, ByVal find As Boolean, ByVal update As Boolean)

            If add = True Then
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            Else
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            End If

            If find = True Then
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            Else
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            End If

            If update Then
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            Else
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            End If

        End Sub

        Public Function GetDocnum_BaseonSeries(ByVal objectcode As String, ByVal Selected_seriescode As String)
            Try
                Dim strsql As String = "Select ""NextNumber"" from nnm1 where ""ObjectCode""='" & objectcode.ToString & "' and ""Series""='" & Selected_seriescode.ToString & "'"
                objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objrs.DoQuery(strsql)
                If objrs.RecordCount = 0 Then Return ""
                Return objrs.Fields.Item(0).Value.ToString
            Catch ex As Exception
                Return ""
            End Try
        End Function

        Public Sub ChooseFromList_Before(ByVal OForm As SAPbouiCOM.Form, ByVal CFLID As String, ByVal SqlQuery_Condition As String, ByVal AliseID As String)
            Dim rsetCFL As SAPbobsCOM.Recordset
            rsetCFL = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Try
                Dim oCFL As SAPbouiCOM.ChooseFromList = OForm.ChooseFromLists.Item(CFLID)
                Dim oConds As SAPbouiCOM.Conditions
                Dim oCond As SAPbouiCOM.Condition
                Dim oEmptyConds As New SAPbouiCOM.Conditions
                oCFL.SetConditions(oEmptyConds)
                oConds = oCFL.GetConditions()
                rsetCFL = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                rsetCFL.DoQuery(SqlQuery_Condition)
                rsetCFL.MoveFirst()
                If rsetCFL.RecordCount > 0 Then
                    For i As Integer = 1 To rsetCFL.RecordCount
                        If i = (rsetCFL.RecordCount) Then
                            oCond = oConds.Add()
                            oCond.Alias = AliseID
                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                            oCond.CondVal = Trim(rsetCFL.Fields.Item(0).Value)
                        Else
                            oCond = oConds.Add()
                            oCond.Alias = AliseID
                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                            oCond.CondVal = Trim(rsetCFL.Fields.Item(0).Value)
                            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                        End If
                        rsetCFL.MoveNext()
                    Next
                Else
                    oCFL.SetConditions(oEmptyConds)
                    oConds = oCFL.GetConditions()
                    oCond = oConds.Add()
                    oCond.Alias = AliseID
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_NONE
                    oCond.CondVal = "-1"
                End If

                oCFL.SetConditions(oConds)
            Catch ex As Exception

            Finally
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rsetCFL)
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        End Sub
        Public Function GetDateTimeValue(ByVal SBODaMIPLAGNTMASring As String) As DateTime
            Dim objBridge As SAPbobsCOM.SBObob
            objBridge = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
            objBridge.Format_StringToDate("")
            Return objBridge.Format_StringToDate(SBODaMIPLAGNTMASring).Fields.Item(0).Value
        End Function
        Public Function getSingleValue(ByVal StrSQL As String) As String
            Try
                Dim rset As SAPbobsCOM.Recordset = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim strReturnVal As String = ""
                rset.DoQuery(StrSQL)
                Return IIf(rset.RecordCount > 0, rset.Fields.Item(0).Value.ToString(), "")
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(" Get Single Value Function Failed :  " & ex.Message + StrSQL, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Return ""
            End Try
        End Function
        Public Function GetSeries(ByVal Objcode As String, ByVal DocDate As String) As String
            Dim series As String = "", Indicator As String

            Indicator = getSingleValue("select ""Indicator""  from OFPR where '" & CDate(DocDate.ToString).ToString("yyyy-MM-dd") & "' between ""F_RefDate"" and ""T_RefDate""")
            If Objcode = "23" Then
                series = getSingleValue("select ""Series"" From  NNM1 where ""ObjectCode""='" & Objcode & "' and ""Indicator""='" & Indicator & "'")
            End If
            If series <> "" Then
                Return series
            Else
                Return ""
            End If
        End Function
        Public Sub RemoveLastrow(ByVal omatrix As SAPbouiCOM.Matrix, ByVal Columname_check As String)
            Try
                If omatrix.VisualRowCount = 0 Then Exit Sub
                If Columname_check.ToString = "" Then Exit Sub
                If omatrix.Columns.Item(Columname_check).Cells.Item(omatrix.VisualRowCount).Specific.string = "" Then
                    omatrix.DeleteRow(omatrix.VisualRowCount)
                End If
            Catch ex As Exception

            End Try
        End Sub
        Public Sub SetCellEdit(ByVal Matrix0 As SAPbouiCOM.Matrix, ByVal EditFlag As Boolean)
            Matrix0.CommonSetting.SetCellEditable(Matrix0.VisualRowCount, 1, EditFlag)
            Matrix0.CommonSetting.SetCellEditable(Matrix0.VisualRowCount, 3, EditFlag)
            Matrix0.CommonSetting.SetCellEditable(Matrix0.VisualRowCount, 5, EditFlag)
            Matrix0.CommonSetting.SetCellEditable(Matrix0.VisualRowCount, 7, EditFlag)
        End Sub

        Public Sub SubConButtonEnabling(ByVal objform As SAPbouiCOM.Form)
            Try
                Dim StrQuery, POItem As String
                Dim objRecSet As SAPbobsCOM.Recordset
                objRecSet = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If objform.Items.Item("6").Specific.String <> "" Then
                    If objaddon.HANA Then
                        StrQuery = "update T0 set ""U_SubConBOM""='Y' From OWOR T0 inner join OITT T1 on T0.""ItemCode""=T1.""Code"" where T1.""U_SubConBOM""='Y'"
                    Else
                        StrQuery = "update T0 set U_SubConBOM='Y' From OWOR T0 inner join OITT T1 on T0.ItemCode=T1.Code where T1.U_SubConBOM='Y'"
                    End If
                    objRecSet.DoQuery(StrQuery)
                End If
                If objaddon.HANA Then
                    POItem = objaddon.objglobalmethods.getSingleValue("select ""U_SubConBOM"" from OITT where ""Code""='" & objform.Items.Item("6").Specific.String & "'")
                Else
                    POItem = objaddon.objglobalmethods.getSingleValue("select U_SubConBOM from OITT where Code='" & objform.Items.Item("6").Specific.String & "'")
                End If
                If POItem = "Y" Then
                    objform.Items.Item("BtnSubCon").Enabled = True
                Else
                    objform.Items.Item("BtnSubCon").Enabled = False
                End If

            Catch ex As Exception
            End Try
        End Sub

        Public Sub LoadSeries(ByVal objform As SAPbouiCOM.Form, ByVal DBSource As SAPbouiCOM.DBDataSource)
            Try
                Dim ComboBox0 As SAPbouiCOM.ComboBox
                ComboBox0 = objform.Items.Item("Series").Specific
                ComboBox0.ValidValues.LoadSeries("SUBPO", SAPbouiCOM.BoSeriesMode.sf_Add)
                ComboBox0.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                DBSource.SetValue("DocNum", 0, objaddon.objglobalmethods.GetDocNum("SUBPO", CInt(ComboBox0.Selected.Value)))
            Catch ex As Exception

            End Try
        End Sub

        Public Function AutoAssign_SubItem(ByVal FolderID As SAPbouiCOM.Folder, ByVal MatrixID As SAPbouiCOM.Matrix) As Boolean
            Dim RowFlag As Boolean = False
            Try
                If MatrixID.VisualRowCount = 1 Then
                    RowFlag = True
                End If
                If FolderID.Item.UniqueID = "flroutput" Then  'Output Folder
                    'If Matrix3.VisualRowCount = 0 Then GoTo addnewrow
                    If MatrixID.Columns.Item("Code").Cells.Item(MatrixID.VisualRowCount).Specific.String <> "" Then
                        If MatrixID.Columns.Item("GRNo").Cells.Item(MatrixID.VisualRowCount).Specific.String <> "" Then
                            RowFlag = True
                        End If
                    End If
                ElseIf FolderID.Item.UniqueID = "flrscrap" Then  'Scrap Folder
                    Dim objCombo As SAPbouiCOM.ComboBox
                    'If Matrix3.VisualRowCount = 0 Then objaddon.objglobalmethods.Matrix_Addrow(Matrix2, "Code", "#")
                    objCombo = MatrixID.Columns.Item("Type").Cells.Item(MatrixID.VisualRowCount).Specific
                    If objCombo.Selected.Description = "Scrap" Then
                        If MatrixID.Columns.Item("GRNo").Cells.Item(MatrixID.VisualRowCount).Specific.String <> "" Then
                            RowFlag = True
                        End If
                    ElseIf objCombo.Selected.Description = "Return" Then
                        If MatrixID.Columns.Item("InvNum").Cells.Item(MatrixID.VisualRowCount).Specific.String <> "" Then
                            RowFlag = True
                            'objaddon.objglobalmethods.Matrix_Addrow(Matrix2, "Code", "#")
                        End If
                    End If
                End If
                Return RowFlag
            Catch ex As Exception
                Return False
            End Try
        End Function

        Public Sub GetSubItemQty(ByVal objform As SAPbouiCOM.Form, ByVal MatrixID As SAPbouiCOM.Matrix)
            Try
                Dim Qty As Double = 0, InvQty As Double = 0
                Dim objRs As SAPbobsCOM.Recordset
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                For i As Integer = 1 To MatrixID.VisualRowCount
                    If MatrixID.Columns.Item("Code").Cells.Item(i).Specific.string <> "" And MatrixID.Columns.Item("Code").Cells.Item(i).Specific.string = objform.Items.Item("txtsitem").Specific.String Then
                        Qty += CDbl(MatrixID.Columns.Item("Quant").Cells.Item(i).Specific.string)
                    End If
                Next
                If objform.Items.Item("txtentry").Specific.String <> "" Then
                    If objaddon.HANA Then
                        objRs.DoQuery("select Top 1 T1.""U_Itemcode"",sum(T1.""U_Qty"") ""Qty"" from ""@MIPL_OPOR"" T0 join ""@MIPL_POR3"" T1 on T0.""DocEntry""=T1.""DocEntry"" " &
                                      "where T1.""U_Type""='2' and T0.""DocEntry""='" & objform.Items.Item("txtentry").Specific.String & "' and T1.""U_Status""='C' and ifnull(T1.""U_InvNo"",'') <>'' and ifnull(T1.""U_GRNo"",'') ='' " &
                                      " group by T1.""U_Itemcode"" order by sum(T1.""U_Qty"") desc")
                    Else
                        objRs.DoQuery("select Top 1 T1.U_Itemcode,sum(T1.U_Qty) Qty from [@MIPL_OPOR] T0 join [@MIPL_POR3] T1 on T0.DocEntry=T1.DocEntry " &
                                      "where T1.U_Type='2' and T0.DocEntry='" & objform.Items.Item("txtentry").Specific.String & "' and T1.U_Status='C' and isnull(T1.U_InvNo,'') <>'' and isnull(T1.U_GRNo,'') ='' " &
                                      " group by T1.U_Itemcode order by sum(T1.U_Qty) desc")
                    End If
                    InvQty = CDbl(objRs.Fields.Item("Qty").Value.ToString)
                End If

                If Qty < CDbl(objform.Items.Item("SQty").Specific.String) Then
                    'Dim cc As String = objform.Items.Item("txtsitem").Specific.String
                    MatrixID.Columns.Item("Code").Cells.Item(MatrixID.VisualRowCount).Specific.string = objform.Items.Item("txtsitem").Specific.String
                    If objaddon.HANA Then
                        MatrixID.Columns.Item("Desc").Cells.Item(MatrixID.VisualRowCount).Specific.String = objaddon.objglobalmethods.getSingleValue("select ""ItemName"" from  OITM where ""ItemCode""='" & MatrixID.Columns.Item("Code").Cells.Item(MatrixID.VisualRowCount).Specific.string & "'")
                    Else
                        MatrixID.Columns.Item("Desc").Cells.Item(MatrixID.VisualRowCount).Specific.String = objaddon.objglobalmethods.getSingleValue("select ItemName from  OITM where ItemCode='" & MatrixID.Columns.Item("Code").Cells.Item(MatrixID.VisualRowCount).Specific.string & "'")
                    End If
                    Dim txtDate As SAPbouiCOM.EditText
                    txtDate = MatrixID.Columns.Item("Date").Cells.Item(MatrixID.VisualRowCount).Specific
                    txtDate.Value = Now.Date.ToString("yyyyMMdd")
                    objform.Update()
                    MatrixID.Columns.Item("Colstat").Cells.Item(MatrixID.VisualRowCount).Specific.String = "O"
                    If CDbl(MatrixID.Columns.Item("Quant").Cells.Item(MatrixID.VisualRowCount).Specific.string) = CDbl(0) Then
                        MatrixID.Columns.Item("Quant").Cells.Item(MatrixID.VisualRowCount).Specific.string = CDbl((objform.Items.Item("SQty").Specific.String) - CDbl(CDbl(Qty) + InvQty))
                    End If

                    Try
                        If objaddon.HANA Then
                            MatrixID.Columns.Item("InvUom").Cells.Item(MatrixID.VisualRowCount).Specific.String = objaddon.objglobalmethods.getSingleValue("select ""InvntryUom"" from  OITM where ""ItemCode""='" & MatrixID.Columns.Item("Code").Cells.Item(MatrixID.VisualRowCount).Specific.String & "'")
                        Else
                            MatrixID.Columns.Item("InvUom").Cells.Item(MatrixID.VisualRowCount).Specific.String = objaddon.objglobalmethods.getSingleValue("select InvntryUom from  OITM where ItemCode='" & MatrixID.Columns.Item("Code").Cells.Item(MatrixID.VisualRowCount).Specific.String & "'")
                        End If
                    Catch ex As Exception
                    End Try
                    If objaddon.HANA Then
                        MatrixID.Columns.Item("HSN").Cells.Item(MatrixID.VisualRowCount).Specific.String = objaddon.objglobalmethods.getSingleValue("Select T0.""ChapterID"" from OCHP T0 join OITM T1 on T0.""AbsEntry""=T1.""ChapterID"" where T1.""ItemCode""='" & MatrixID.Columns.Item("Code").Cells.Item(MatrixID.VisualRowCount).Specific.String & "'")
                    Else
                        MatrixID.Columns.Item("HSN").Cells.Item(MatrixID.VisualRowCount).Specific.String = objaddon.objglobalmethods.getSingleValue("Select T0.ChapterID from OCHP T0 join OITM T1 on T0.AbsEntry=T1.ChapterID where T1.ItemCode='" & MatrixID.Columns.Item("Code").Cells.Item(MatrixID.VisualRowCount).Specific.String & "'")

                    End If
                    'If ToLocationInOutput = "Y" Then
                    '    If objaddon.HANA Then
                    '        MatrixID.Columns.Item("Whse").Cells.Item(MatrixID.VisualRowCount).Specific.string = objaddon.objglobalmethods.getSingleValue("Select DflWhs"" from OBPL where ""BPLId""=(select  distinct T0.""BPLid"" from OWHS T0 left join OLCT T1 on T0.""Location""=T1.""Code"" where T1.""Code""=(Select ""U_LCodeO"" from ""@MIPL_GEN"") and ifnull(T0.""Inactive"",'') ='N' and T0.""BPLid"" is not null)")
                    '    Else
                    '        MatrixID.Columns.Item("Whse").Cells.Item(MatrixID.VisualRowCount).Specific.string = objaddon.objglobalmethods.getSingleValue("Select DflWhs from OBPL where BPLId=(select  distinct T0.BPLid from OWHS T0 left join OLCT T1 on T0.Location=T1.Code where T1.Code=(Select U_LCodeO from [@MIPL_GEN]) and isnull(T0.Inactive,'') ='N' and T0.BPLid is not null)")
                    '    End If
                    'End If
                    objform.Update()
                Else
                    Exit Sub
                End If
            Catch ex As Exception

            End Try
        End Sub
    End Class

End Namespace
