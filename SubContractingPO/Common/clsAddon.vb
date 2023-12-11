Imports SAPbouiCOM.Framework
Imports System.IO


Namespace SubContractingPO
    Public Class clsAddon
        Public WithEvents objapplication As SAPbouiCOM.Application
        Public objcompany As SAPbobsCOM.Company
        Public objmenuevent As clsMenuEvent
        Public objrightclickevent As clsRightClickEvent
        Public objglobalmethods As clsGlobalMethods
        Dim objform As SAPbouiCOM.Form
        Dim strsql As String = ""
        Dim objrs As SAPbobsCOM.Recordset
        Dim print_close As Boolean = False
        Public HANA As Boolean = True
        'Public HANA As Boolean = False
        Public HWKEY() As String = New String() {"V0913316776", "T0897075900", "Y2146381495", "L1653539483", "R1574408489", "J1776539436", "K1679825911", "B1133561807", "M0090876837", "Y1334940735", "A0061802481", "T0659980171"}
        Public Sub Intialize(ByVal args() As String)
            Try
                Dim oapplication As Application
                If (args.Length < 1) Then oapplication = New Application Else oapplication = New Application(args(0))
                objapplication = Application.SBO_Application
                If isValidLicense() Then
                    objapplication.StatusBar.SetText("Establishing Company Connection Please Wait...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    objcompany = Application.SBO_Application.Company.GetDICompany()

                    Create_DatabaseFields() 'UDF & UDO Creation Part    
                    Menu() 'Menu Creation Part
                    Create_Objects() 'Object Creation Part

                    objapplication.StatusBar.SetText("Addon Connected Successfully..!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    oapplication.Run()
                Else
                    objapplication.StatusBar.SetText("Addon Disconnected due to license mismatch..!!!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Sub
                End If
                'System.Windows.Forms.Application.Run()
            Catch ex As Exception
                objaddon.objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End Try
        End Sub
        Function isValidLicense() As Boolean
            Try
                Try
                    If objapplication.Forms.ActiveForm.TypeCount > 0 Then
                        For i As Integer = 0 To objapplication.Forms.ActiveForm.TypeCount - 1
                            objapplication.Forms.ActiveForm.Close()
                        Next
                    End If
                Catch ex As Exception
                End Try

                'If Not HANA Then
                '    objapplication.Menus.Item("1030").Activate()
                'End If
                objapplication.Menus.Item("257").Activate()
                Dim CrrHWKEY As String = objapplication.Forms.ActiveForm.Items.Item("79").Specific.Value.ToString.Trim
                objapplication.Forms.ActiveForm.Close()

                For i As Integer = 0 To HWKEY.Length - 1
                    If HWKEY(i).Trim = CrrHWKEY.Trim Then
                        Return True
                    End If
                Next
                MsgBox("Installing Add-On failed due to License mismatch", MsgBoxStyle.OkOnly, "License Management")
                Return False
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                'MsgBox(ex.ToString)
            End Try
            Return True
        End Function
        Private Sub Create_Objects()
            objmenuevent = New clsMenuEvent
            objrightclickevent = New clsRightClickEvent
            objglobalmethods = New clsGlobalMethods
        End Sub

        Private Sub Create_DatabaseFields()
            'If objapplication.Company.UserName.ToString.ToUpper <> "MANAGER" Then

            'If objapplication.MessageBox("Do you want to execute the field Creations?", 2, "Yes", "No") <> 1 Then Exit Sub
            objapplication.StatusBar.SetText("Creating Database Fields.Please Wait...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            Dim objtable As New clsTable
            objtable.FieldCreation()
            'End If

        End Sub


#Region "Menu Creation Details"

        Private Sub Menu()
            Dim Menucount As Integer = 1
            CreateMenu("", Menucount, "Sub-Contracting", SAPbouiCOM.BoMenuType.mt_POPUP, "SUBPO", "2304")  '4352
            Menucount = 1 'Menu Inside  

            CreateMenu("", Menucount, "Sub-Con BOM", SAPbouiCOM.BoMenuType.mt_STRING, "SUBCT", "SUBPO") : Menucount += 1
            CreateMenu("", Menucount, "Sub Contracting", SAPbouiCOM.BoMenuType.mt_STRING, "SUBCTPO", "SUBPO") : Menucount += 1
            Menucount = 12
            CreateMenu("", Menucount, "Sub-Contracting", SAPbouiCOM.BoMenuType.mt_POPUP, "SUBPOG", "43525")
            Menucount = 1 'Menu Inside   
            CreateMenu("", Menucount, "Settings", SAPbouiCOM.BoMenuType.mt_STRING, "SUBGEN", "SUBPOG") : Menucount += 1

        End Sub

        Private Sub CreateMenu(ByVal ImagePath As String, ByVal Position As Int32, ByVal DisplayName As String, ByVal MenuType As SAPbouiCOM.BoMenuType, ByVal UniqueID As String, ByVal ParentMenuID As String)
            Try
                Dim oMenuPackage As SAPbouiCOM.MenuCreationParams
                Dim parentmenu As SAPbouiCOM.MenuItem
                parentmenu = objapplication.Menus.Item(ParentMenuID)
                If parentmenu.SubMenus.Exists(UniqueID.ToString) Then Exit Sub
                oMenuPackage = objapplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                oMenuPackage.Image = ImagePath
                oMenuPackage.Position = Position
                oMenuPackage.Type = MenuType
                oMenuPackage.UniqueID = UniqueID
                oMenuPackage.String = DisplayName
                parentmenu.SubMenus.AddEx(oMenuPackage)
            Catch ex As Exception
                objapplication.StatusBar.SetText("Menu Already Exists", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
            End Try
            'Return ParentMenu.SubMenus.Item(UniqueID)
        End Sub

#End Region

#Region "ItemEvent_Link Button"

        Private Sub objapplication_ItemEvent(FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles objapplication.ItemEvent
            Try
                If pVal.BeforeAction Then
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_CLICK
                            If bModal And (objaddon.objapplication.Forms.ActiveForm.TypeEx = "SUBCTPO") Then
                                BubbleEvent = False
                                Try
                                    If objapplication.Forms.Item("ViewD").VisibleEx Then
                                        objapplication.Forms.Item("ViewD").Select()
                                    End If
                                Catch ex As Exception
                                End Try


                            End If
                        Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
                            Dim EventEnum As SAPbouiCOM.BoEventTypes
                            EventEnum = pVal.EventType
                            If FormUID = "ViewD" And (EventEnum = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE) And bModal Then
                                bModal = False
                            End If

                    End Select
                Else
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE

                    End Select
                End If

            Catch ex As Exception

            End Try
        End Sub



#End Region

#Region "Menu Event"

        Public Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles objapplication.MenuEvent
            Try
                Select Case pVal.MenuUID
                    Case "1281", "1282", "1283", "1284", "1285", "1286", "1287", "1300", "1288", "1289", "1290", "1291", "1304", "1292", "1293", "DUP", "6913"
                        objmenuevent.MenuEvent_For_StandardMenu(pVal, BubbleEvent)
                    Case "SUBCTPO", "SUBCT", "SUBGEN"
                        MenuEvent_For_FormOpening(pVal, BubbleEvent)
                        'Case "1293"
                        '    BubbleEvent = False
                    Case "519"
                        MenuEvent_For_Preview(pVal, BubbleEvent)
                End Select
            Catch ex As Exception
                'objaddon.objapplication.SetStatusBarMessage("Error in SBO_Application MenuEvent" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try
        End Sub

        Public Sub MenuEvent_For_Preview(ByRef pval As SAPbouiCOM.MenuEvent, ByRef bubbleevent As Boolean)
            Dim oform = objaddon.objapplication.Forms.ActiveForm()
            'If pval.BeforeAction Then
            '    If oform.TypeEx = "TRANOLVA" Then MenuEvent_For_PrintPreview(oform, "8f481d5cf08e494f9a83e1e46ab2299e", "txtentry") : bubbleevent = False
            '    If oform.TypeEx = "TRANOLAP" Then MenuEvent_For_PrintPreview(oform, "f15ee526ac514070a9d546cda7f94daf", "txtentry") : bubbleevent = False
            '    If oform.TypeEx = "OLSE" Then MenuEvent_For_PrintPreview(oform, "e47ed373e0cc48efb47c9773fba64fc3", "txtentry") : bubbleevent = False
            'End If
        End Sub

        Private Sub MenuEvent_For_PrintPreview(ByVal oform As SAPbouiCOM.Form, ByVal Menuid As String, ByVal Docentry_field As String)
            'Try
            '    Dim Docentry_Est As String = oform.Items.Item(Docentry_field).Specific.String
            '    If Docentry_Est = "" Then Exit Sub
            '    print_close = False
            '    objaddon.objapplication.Menus.Item(Menuid).Activate()
            '    oform = objaddon.objapplication.Forms.ActiveForm()
            '    oform.Items.Item("1000003").Specific.string = Docentry_Est
            '    oform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            '    print_close = True
            'Catch ex As Exception
            'End Try
        End Sub
        Public Function FormExist(ByVal FormID As String) As Boolean
            Try
                FormExist = False
                For Each uid As SAPbouiCOM.Form In objaddon.objapplication.Forms
                    If uid.UniqueID = FormID Then
                        FormExist = True
                        Exit For
                    End If
                Next
                If FormExist Then
                    If FormID = "BOM" Then
                        Try
                            Dim cflForm As SAPbouiCOM.Form
                            If objaddon.objapplication.Forms.Count > 0 Then
                                For frm As Integer = 0 To objaddon.objapplication.Forms.Count - 1
                                    If objaddon.objapplication.Forms.Item(frm).UniqueID = "BOM" Then
                                        cflForm = objaddon.objapplication.Forms.Item("BOM")
                                        cflForm.Close()
                                        Return False
                                        Exit For
                                    End If
                                Next
                            End If
                        Catch ex As Exception
                        End Try
                    Else
                        objaddon.objapplication.Forms.Item(FormID).Visible = True
                        objaddon.objapplication.Forms.Item(FormID).Select()
                    End If

                End If
            Catch ex As Exception

            End Try

        End Function

        Public Sub MenuEvent_For_FormOpening(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Try
                If pVal.BeforeAction = False Then
                    Select Case pVal.MenuUID
                        Case "SUBCTPO"
                            NewLink = "New"
                            If Not FormExist("SUBCTPO") Then
                                Dim activeform As New SubContractingPO
                                activeform.Show()
                            End If

                            NewLink = "-1"
                        Case "SUBCT"
                            'If Not FormExist("BOM") Then
                            '    Dim activeform As New SubContract
                            '    activeform.Show()
                            'End If
                            Dim activeform As New SubContract
                            activeform.Show()
                        Case "SUBGEN"
                            If Not FormExist("SubGen") Then
                                Dim activeform As New GeneralSettings
                                activeform.Show()
                            End If
                            
                    End Select

                End If
            Catch ex As Exception
                'objaddon.objapplication.SetStatusBarMessage("Error in Form Opening MenuEvent" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try
        End Sub

#End Region

#Region "LayoutKeyEvent"

        Public Sub SBO_Application_LayoutKeyEvent(ByRef eventInfo As SAPbouiCOM.LayoutKeyInfo, ByRef BubbleEvent As Boolean) Handles objapplication.LayoutKeyEvent
            'Dim oForm_Layout As SAPbouiCOM.Form = Nothing
            'If SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.BusinessObject.Type = "NJT_CES" Then
            '    oForm_Layout = SAPbouiCOM.Framework.Application.SBO_Application.Forms.Item(eventInfo.FormUID)
            'End If
        End Sub

#End Region

#Region "Application Event"

        Public Sub SBO_Application_AppEvent(EventType As SAPbouiCOM.BoAppEventTypes) Handles objapplication.AppEvent
            Select Case EventType
                Case SAPbouiCOM.BoAppEventTypes.aet_ShutDown Or SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged Or SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition
                    Try
                        If objcompany.Connected Then objcompany.Disconnect()
                        'System.Runtime.InteropServices.Marshal.ReleaseComObject(objcompany)
                        'System.Runtime.InteropServices.Marshal.ReleaseComObject(objapplication)
                        objcompany = Nothing
                        'objapplication = Nothing
                        GC.Collect()
                        System.Windows.Forms.Application.Exit()
                        End
                    Catch ex As Exception
                    End Try
                    'End
                    'Case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged
                    '    End
                    '    'Case SAPbouiCOM.BoAppEventTypes.aet_FontChanged
                    '    '    End
                    '    'Case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged
                    '    '    End
                    'Case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition
                    '    End
            End Select
        End Sub

#End Region

#Region "Right Click Event"

        Private Sub objapplication_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles objapplication.RightClickEvent
            Try
                Select Case objaddon.objapplication.Forms.ActiveForm.TypeEx
                    Case "SUBCTPO", "SUBBOM", "65211", "SUBGEN"
                        objrightclickevent.RightClickEvent(eventInfo, BubbleEvent)
                End Select
            Catch ex As Exception

            End Try
        End Sub

#End Region


    End Class
End Namespace
