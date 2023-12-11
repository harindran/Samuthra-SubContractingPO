Namespace SubContractingPO

    Public Class clsTable

        Public Sub FieldCreation()
            SubContractingPO()
            SubContractingBOM()
            'Costing()
            GeneralSettings()

            AddFields("OITM", "SubCont", "Sub Contract", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , , , {"Y,Yes", "N,No"})
            'AddFields("OIGN", "TabType", "TabType", SAPbobsCOM.BoFieldTypes.db_Alpha, 7)
            AddFields("OWTR", "SubConNo", "SubContracting Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)

            AddFields("OIGN", "SeyDCNum", "Seyoon DC Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("OIGN", "SupDCNum", "Supplier DC Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("OCRD", "WAREHOUSE", "WAREHOUSE", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("OWTR", "OutNum", "Output Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("OWTR", "ScrapNum", "Scrap Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)

            AddFields("OJDT", "SubConNo", "SubContracting Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("OJDT", "Status", "JE Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 5, , , , , {"O,Open", "C,Close"})
            AddFields("WTR1", "PlanQty", "Planned Quantity", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("IGN1", "PlanQty", "Planned Quantity", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("IGN1", "Price", "Unit Price", SAPbobsCOM.BoFieldTypes.db_Float, 6, SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("OIGE", "SubConNo", "SubContracting Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)   'Goods Issue
            AddFields("OIGN", "SubConNo", "SubContracting Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)  'Goods Receipt
            AddFields("IGN1", "CardCode", "Card Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 15)
            AddFields("IGN1", "LineID", "Line Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 15)
            AddFields("IGN1", "TabType", "TabType", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("OITT", "SubConBOM", "Sub Contract BOM", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , , , {"Y,Yes", "N,No"})
            AddFields("OWOR", "SubConBOM", "Sub Contract BOM", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , , , {"Y,Yes", "N,No"})
            AddFields("OWOR", "SubPONum", "SubContracting Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("OWOR", "CardCode", "SubCon CardCode", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        End Sub



#Region "Document Data Creation"

        Private Sub SubContractingPO()
            AddTables("MIPL_OPOR", "SubContracting PO Header", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddTables("MIPL_POR1", "SubContracting Input Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddTables("MIPL_POR2", "SubContracting Output Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddTables("MIPL_POR3", "SubContracting Scrap Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddTables("MIPL_POR4", "SubContracting RelDoc Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddTables("MIPL_POR5", "SubContracting Costing Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

            AddFields("@MIPL_OPOR", "CardCode", "Vendor Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 200)
            AddFields("@MIPL_OPOR", "InvUom", "Inventory Uom", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MIPL_OPOR", "CardName", "Vendor Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 200)
            AddFields("@MIPL_OPOR", "ContPerson", "Contact Person", SAPbobsCOM.BoFieldTypes.db_Alpha, 200)
            AddFields("@MIPL_OPOR", "VenRefNo", "VendorRef Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 200)
            AddFields("@MIPL_OPOR", "SItemCode", "Sub ItemCode", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MIPL_OPOR", "BOMCode", "BOM ItemCode", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MIPL_OPOR", "SItemDesc", "Sub ItemDesc", SAPbobsCOM.BoFieldTypes.db_Alpha, 200)
            AddFields("@MIPL_OPOR", "VBalnce", "Vendor Balance", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("@MIPL_OPOR", "VOBStock", "Vendor OB Stock", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            AddFields("@MIPL_OPOR", "SQty", "Sub Item Qty", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            AddFields("@MIPL_OPOR", "DocDate", "Document Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("@MIPL_OPOR", "DocDueDate", "Document DueDate", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("@MIPL_OPOR", "TaxDate", "Tax Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("@MIPL_OPOR", "PONum", "PO Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 15)
            AddFields("@MIPL_OPOR", "PurOrdrNo", "PurchaseOrder Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 15)
            AddFields("@MIPL_OPOR", "InvTrNo", "InventoryTransfer Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("@MIPL_OPOR", "OpenQty", "SubItem Open Qty", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            AddFields("@MIPL_OPOR", "GRNo", "GoodsRec Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("@MIPL_OPOR", "GINo", "GoodsIssue Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("@MIPL_OPOR", "MtxEdit", "Matrix Editable", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , , , {"Y,Yes", "N,No"})
            AddFields("@MIPL_OPOR", "POLine", "PO LineNumber", SAPbobsCOM.BoFieldTypes.db_Alpha, 15)
            AddFields("@MIPL_OPOR", "PODoc", "PO DocNumber", SAPbobsCOM.BoFieldTypes.db_Alpha, 15)
            AddFields("@MIPL_OPOR", "PurEnt", "PurChase DocNumber", SAPbobsCOM.BoFieldTypes.db_Alpha, 15)
            AddFields("@MIPL_OPOR", "EditQty", "Editable Qty", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "N", True)
            AddFields("@MIPL_OPOR", "Clstat", "Closing Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "N", True)
            'AddFields("@MIPL_OPOR", "BOMType", "BOM Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            'AddFields("@MIPL_OPOR", "Rounding", "Rounding Amount", SAPbobsCOM.BoFieldTypes.db_Float, 6, SAPbobsCOM.BoFldSubTypes.st_Price)
            'AddFields("@MIPL_OPOR", "BefDisc", "Before Discount Amount", SAPbobsCOM.BoFieldTypes.db_Float, 6, SAPbobsCOM.BoFldSubTypes.st_Price)
            'AddFields("@MIPL_OPOR", "TaxAmount", "Tax Amount", SAPbobsCOM.BoFieldTypes.db_Float, 6, SAPbobsCOM.BoFldSubTypes.st_Price)
            'AddFields("@MIPL_OPOR", "DiscAmount", "Discount Amount", SAPbobsCOM.BoFieldTypes.db_Float, 6, SAPbobsCOM.BoFldSubTypes.st_Price)
            'AddFields("@MIPL_OPOR", "TotAmount", "Total Amount", SAPbobsCOM.BoFieldTypes.db_Float, 6, SAPbobsCOM.BoFldSubTypes.st_Price)
            'AddFields("@MIPL_OPOR", "DiscPer", "Discount Percent", SAPbobsCOM.BoFieldTypes.db_Float, 6, SAPbobsCOM.BoFldSubTypes.st_Price)

            AddFields("@MIPL_POR1", "Itemcode", "ItemCode", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MIPL_POR1", "Item1", "ItemCode 1", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MIPL_POR1", "InvUom", "Inventory Uom", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MIPL_POR1", "ItemDesc", "ItemDesc", SAPbobsCOM.BoFieldTypes.db_Alpha, 200)
            AddFields("@MIPL_POR1", "Qty", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            AddFields("@MIPL_POR1", "PlanQty", "Planned Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            AddFields("@MIPL_POR1", "Price", "Price", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("@MIPL_POR1", "LineTot", "Line Total", SAPbobsCOM.BoFieldTypes.db_Float, 6, SAPbobsCOM.BoFldSubTypes.st_Price)
            'AddFields("@MIPL_POR1", "TaxLine", "Tax LineAmount", SAPbobsCOM.BoFieldTypes.db_Float, 6, SAPbobsCOM.BoFldSubTypes.st_Price)
            'AddFields("@MIPL_POR1", "TaxCode", "Tax Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 40)
            AddFields("@MIPL_POR1", "WhsCode", "Warehouse Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            'AddFields("@MIPL_POR1", "DiscPer", "Discount Percent", SAPbobsCOM.BoFieldTypes.db_Float, 6, SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("@MIPL_POR1", "HSNCode", "HSN Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 20)
            AddFields("@MIPL_POR1", "InStock", "In Stock", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("@MIPL_POR1", "SubWhse", "Sub Whscode", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)

            AddFields("@MIPL_POR1", "ProcQty", "Processed Qty", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            AddFields("@MIPL_POR1", "OpenQty", "Open Qty", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            AddFields("@MIPL_POR1", "RetQty", "Returned Qty", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            'AddFields("@MIPL_POR1", "Process", "Process", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MIPL_POR1", "Weight", "Weight", SAPbobsCOM.BoFieldTypes.db_Float, 6, SAPbobsCOM.BoFldSubTypes.st_Measurement)
            AddFields("@MIPL_POR1", "LTType", "LineTotal Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, , , "1", , {"1,Default", "2,Weight"})
            AddFields("@MIPL_POR1", "ProType", "Process Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, , , "1", , {"1,Process1", "2,Process2"})

            AddFields("@MIPL_POR2", "Itemcode", "ItemCode", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MIPL_POR2", "ItemDesc", "ItemDesc", SAPbobsCOM.BoFieldTypes.db_Alpha, 200)
            AddFields("@MIPL_POR2", "InvUom", "Inventory Uom", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MIPL_POR2", "Qty", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            AddFields("@MIPL_POR2", "Price", "Price", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("@MIPL_POR2", "LineTot", "Line Total", SAPbobsCOM.BoFieldTypes.db_Float, 6, SAPbobsCOM.BoFldSubTypes.st_Price)
            'AddFields("@MIPL_POR2", "GRCheck", "GoodsRec Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 2, , , "N", True)
            AddFields("@MIPL_POR2", "WhsCode", "Warehouse Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MIPL_POR2", "HSNCode", "HSN Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 20)
            AddFields("@MIPL_POR2", "InStock", "In Stock", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("@MIPL_POR2", "TabType", "Tab Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("@MIPL_POR2", "Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, , , , , {"O,Open", "C,Close"})
            'AddFields("@MIPL_POR2", "Date", "Document Date", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("@MIPL_POR2", "Date", "Document Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("@MIPL_POR2", "ProCost", "Process Cost", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("@MIPL_POR2", "TProCost", "Total ProcessCost", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)

            AddFields("@MIPL_POR2", "GRNo", "GoodsRec Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("@MIPL_POR2", "GINo", "Goodsissue Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("@MIPL_POR2", "InvNo", "InvTran Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            'AddFields("@MIPL_POR2", "InvStat", "Inv Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, , , "O", , {"O,Open", "C,Close"})
            AddFields("@MIPL_POR2", "SeyDCNum", "Seyoon DC Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MIPL_POR2", "SupDCNum", "Supplier DC Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MIPL_POR2", "DocDate", "Document Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("@MIPL_POR2", "RefNo", "Reference Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MIPL_POR2", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 200)

            AddFields("@MIPL_POR3", "Itemcode", "ItemCode", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MIPL_POR3", "ItemDesc", "ItemDesc", SAPbobsCOM.BoFieldTypes.db_Alpha, 200)
            AddFields("@MIPL_POR3", "InvUom", "Inventory Uom", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MIPL_POR3", "Qty", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            AddFields("@MIPL_POR3", "Price", "Price", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("@MIPL_POR3", "LineTot", "Line Total", SAPbobsCOM.BoFieldTypes.db_Float, 6, SAPbobsCOM.BoFldSubTypes.st_Price)
            'AddFields("@MIPL_POR3", "GRCheck", "GoodsRec Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 2, , , "N", True)
            AddFields("@MIPL_POR3", "WhsCode", "Warehouse Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MIPL_POR3", "HSNCode", "HSN Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 20)
            AddFields("@MIPL_POR3", "InStock", "In Stock", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("@MIPL_POR3", "TabType", "Tab Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("@MIPL_POR3", "Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, , , , , {"O,Open", "C,Close"})
            'AddFields("@MIPL_POR3", "Date", "Document Date", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("@MIPL_POR3", "DocDate", "Document Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("@MIPL_POR3", "RefNo", "Reference Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MIPL_POR3", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 200)

            AddFields("@MIPL_POR3", "Date", "Document Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("@MIPL_POR3", "GRNo", "GoodsRec Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("@MIPL_POR3", "InvNo", "InvTran Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            'AddFields("@MIPL_POR3", "InvStat", "Inv Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, , , "O", , {"O,Open", "C,Close"})
            AddFields("@MIPL_POR3", "Type", "Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 7, , , "1", , {"1,Scrap", "2,Return"})


            AddFields("@MIPL_POR4", "DocDate", "Document Date", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("@MIPL_POR4", "Itemcode", "ItemCode", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MIPL_POR4", "ItemDesc", "ItemDesc", SAPbobsCOM.BoFieldTypes.db_Alpha, 200)
            AddFields("@MIPL_POR4", "Qty", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            AddFields("@MIPL_POR4", "WhsCode", "Warehouse Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MIPL_POR4", "LineTot", "Line Total", SAPbobsCOM.BoFieldTypes.db_Float, 6, SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("@MIPL_POR4", "DocNum", "Document Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("@MIPL_POR4", "DocType", "Document Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 10, , , "I", , {"I,Item", "S,Service"})
            AddFields("@MIPL_POR4", "DocDet", "Document Details", SAPbobsCOM.BoFieldTypes.db_Alpha, 10, , , "13", , {"13,A/R Invoice", "18,A/P Invoice"})

            AddFields("@MIPL_POR5", "DocDate", "Document Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("@MIPL_POR5", "GLCode", "GLCode", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MIPL_POR5", "GLName", "GLName", SAPbobsCOM.BoFieldTypes.db_Alpha, 200)
            AddFields("@MIPL_POR5", "Debit", "Debit Amount", SAPbobsCOM.BoFieldTypes.db_Float, 6, SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("@MIPL_POR5", "Credit", "Credit Amount", SAPbobsCOM.BoFieldTypes.db_Float, 6, SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("@MIPL_POR5", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MIPL_POR5", "CostCent", "Cost Center", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MIPL_POR5", "CostCent1", "Cost Center", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MIPL_POR5", "CostCent2", "Cost Center", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MIPL_POR5", "CostCent3", "Cost Center", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MIPL_POR5", "CostCent4", "Cost Center", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MIPL_POR5", "Project", "Project", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MIPL_POR5", "Branch", "Branch Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 40)
            AddFields("@MIPL_POR5", "Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, , , , , {"O,Open", "C,Close"})
            AddFields("@MIPL_POR5", "JENo", "JE Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)

            AddUDO("SUBPO", "SubContractingPO", SAPbobsCOM.BoUDOObjType.boud_Document, "MIPL_OPOR", {"MIPL_POR1", "MIPL_POR2", "MIPL_POR3", "MIPL_POR4", "MIPL_POR5"}, {"DocEntry", "DocNum", "U_SItemCode", "U_CardCode", "U_DocDate", "U_DocDueDate", "U_TaxDate", "U_VenRefNo", "U_PONum", "U_InvTrNo", "U_GRNo", "U_GINo"}, True, True)
        End Sub


#End Region

#Region "Master Data Creation"

        Private Sub GeneralSettings()
            AddTables("MIPL_GEN", "SubContracting General", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            
            AddFields("@MIPL_GEN", "ResEn", "Resources Enable", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "N", True)
            AddFields("@MIPL_GEN", "DatePO", "Date PO", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "N", True)
            AddFields("@MIPL_GEN", "ItemBOM", "ItemCode BOM", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "N", True)
            AddFields("@MIPL_GEN", "Costing", "Costing", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "N", True)
            AddFields("@MIPL_GEN", "POItem", "PO Item", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "N", True)
            AddFields("@MIPL_GEN", "CopBOM", "CopyTo Sub-BOM", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "N", True)
            AddFields("@MIPL_GEN", "AutoItem", "Auto ItemInOutput", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "N", True)
            AddFields("@MIPL_GEN", "Type", "Type InScrap", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "N", True)

            AddFields("@MIPL_GEN", "AutoPO", "Auto ProdOrder", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "N", True)
            AddFields("@MIPL_GEN", "RecLoad", "Receipt Load", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "N", True)
            AddFields("@MIPL_GEN", "APLoad", "APInvoice Load", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "N", True)
            AddFields("@MIPL_GEN", "SubScreen", "SubPO Screen", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "N", True)
            AddFields("@MIPL_GEN", "ScrapCon", "Scrap CFL Condition", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "N", True)
            AddFields("@MIPL_GEN", "SGroup", "Scrap Group", SAPbobsCOM.BoFieldTypes.db_Alpha, 12)

            AddFields("@MIPL_GEN", "ToWhse", "To Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "N", True)
            AddFields("@MIPL_GEN", "Price", "Vendor SplPrice", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "N", True)
            AddFields("@MIPL_GEN", "LCode", "Location Code in Input", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("@MIPL_GEN", "LCodeO", "Location Code Output", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("@MIPL_GEN", "ToWhseO", "To Warehouse In Output", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "N", True)

            AddFields("@MIPL_GEN", "GICode", "GoodsIssue Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MIPL_GEN", "GRCodeO", "GoodsReceipt CodeO", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MIPL_GEN", "GRCodeS", "GoodsReceipt CodeS", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            'AddFields("@MIPL_GEN", "JECode", "Journal Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MIPL_GEN", "GIName", "GoodsIssue Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("@MIPL_GEN", "GRNameO", "GoodsReceipt NameO", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("@MIPL_GEN", "GRNameS", "GoodsReceipt NameS", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            'AddFields("@MIPL_GEN", "JEName", "Journal Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("@MIPL_GEN", "InvWhse", "ToWhse InvTransfer", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "N", True)
            AddFields("@MIPL_GEN", "InvWCode", "ToWhse Code InvTran", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MIPL_GEN", "TranList", "Tran List", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "N", True)
            AddFields("@MIPL_GEN", "SUser", "Super User", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("@MIPL_GEN", "WPrice", "Weight based Price", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "N", True)
            AddFields("@MIPL_GEN", "Field1", "Field 1", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "N", True)
            AddFields("@MIPL_GEN", "UDF0", "UDF 0", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "N", True)
            AddFields("@MIPL_GEN", "UDF1", "UDF 1", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "N", True)
            AddFields("@MIPL_GEN", "Val0", "Validation 0", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "N", True)
            AddFields("@MIPL_GEN", "Val1", "Validation 1", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "N", True)
            'AddFields("@MIPL_GEN", "Val0", "Validation 0", SAPbobsCOM.BoFieldTypes.db_Alpha, 1)
            'AddFields("@MIPL_GEN", "Val1", "Validation 1", SAPbobsCOM.BoFieldTypes.db_Alpha, 1)
            'AddFields("@MIPL_GEN", "Val0", "Validation 0", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, , , "N", , {"Y,Yes", "N,No"})
            'AddFields("@MIPL_GEN", "Val1", "Validation 1", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, , , "N", , {"Y,Yes", "N,No"})
            AddUDO("GENSET", "Sub-Con General Settings", SAPbobsCOM.BoUDOObjType.boud_MasterData, "MIPL_GEN", {""}, {"Code", "Name"}, True, False)
        End Sub

        Private Sub SubContractingBOM()
            AddTables("MIPL_OBOM", "SubContracting BOM Header", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            AddTables("MIPL_BOM1", "SubContracting BOM Lines", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)

            AddFields("@MIPL_OBOM", "ItemCode", "BOM ItemCode", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MIPL_OBOM", "DocEntry", "Document Entry", SAPbobsCOM.BoFieldTypes.db_Alpha, 12)
            AddFields("@MIPL_OBOM", "Qty", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            AddFields("@MIPL_OBOM", "BOMType", "BOM Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MIPL_OBOM", "Type", "Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, , , "-", , {"-,-", "1,Process1", "2,Process2"})
            AddFields("@MIPL_OBOM", "WhseCode", "Warehouse Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MIPL_OBOM", "Project", "Project Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MIPL_OBOM", "Avgplan", "Average Plan", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            AddFields("@MIPL_OBOM", "Distrule", "Distribution Rule", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)

            AddFields("@MIPL_BOM1", "Itemcode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MIPL_BOM1", "ItemDesc", "Item Description", SAPbobsCOM.BoFieldTypes.db_Alpha, 200)
            AddFields("@MIPL_BOM1", "Type", "Item Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 20)
            AddFields("@MIPL_BOM1", "IType", "IType", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, , , "4", , {"4,Item", "290,Resource"})
            AddFields("@MIPL_BOM1", "Qty", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            AddFields("@MIPL_BOM1", "UOMName", "UOM Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 20)
            AddFields("@MIPL_BOM1", "Whse", "Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MIPL_BOM1", "Distrule", "Distribution Rule", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MIPL_BOM1", "Unitprice", "Unit Price", SAPbobsCOM.BoFieldTypes.db_Float, 6, SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("@MIPL_BOM1", "Total", "Total", SAPbobsCOM.BoFieldTypes.db_Float, 6, SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("@MIPL_BOM1", "Comments", "Comments", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)

            AddUDO("SUBBOM", "SubContractingBOM", SAPbobsCOM.BoUDOObjType.boud_MasterData, "MIPL_OBOM", {"MIPL_BOM1"}, {"Code", "Name", "U_DocEntry", "U_Qty", "U_BOMType", "U_WhseCode"}, True, False)
        End Sub

        Private Sub Costing()
            AddTables("MIPL_SBGL", "Sub-GL", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            AddFields("@MIPL_SBGL", "ItmGrp", "Sub ItemGroup", SAPbobsCOM.BoFieldTypes.db_Alpha, 20)
            AddFields("@MIPL_SBGL", "InvWhse", "Inventory Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MIPL_SBGL", "WhsCode", "Sub Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("@MIPL_SBGL", "GoodIssue", "Sub GoodsIssue", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("@MIPL_SBGL", "GoodsReceipt", "Sub GoodReceipt", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("@MIPL_SBGL", "JEGLCode", "JE GLCode", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("@MIPL_SBGL", "JEGLName", "JE GLName", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MIPL_SBGL", "BranchID", "Branch ID", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("@MIPL_SBGL", "BranchNam", "Branch Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 40)
            AddUDO("SUBGL", "SubContractingGL", SAPbobsCOM.BoUDOObjType.boud_MasterData, "MIPL_SBGL", {""}, {"Code", "Name", "U_WhsCode"}, True, False)
        End Sub
#End Region

#Region "Table Creation Common Functions"

        Private Sub AddTables(ByVal strTab As String, ByVal strDesc As String, ByVal nType As SAPbobsCOM.BoUTBTableType)
            Dim oUserTablesMD As SAPbobsCOM.UserTablesMD
            Try
                oUserTablesMD = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
                'Adding Table
                If Not oUserTablesMD.GetByKey(strTab) Then
                    oUserTablesMD.TableName = strTab
                    oUserTablesMD.TableDescription = strDesc
                    oUserTablesMD.TableType = nType

                    If oUserTablesMD.Add <> 0 Then
                        Throw New Exception(objaddon.objcompany.GetLastErrorDescription & strTab)
                    End If
                End If
            Catch ex As Exception
                Throw ex
            Finally
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD)
                oUserTablesMD = Nothing
                GC.WaitForPendingFinalizers()
                GC.Collect()
            End Try
        End Sub

        Private Sub AddFields(ByVal strTab As String, ByVal strCol As String, ByVal strDesc As String, ByVal nType As SAPbobsCOM.BoFieldTypes, _
                             Optional ByVal nEditSize As Integer = 10, Optional ByVal nSubType As SAPbobsCOM.BoFldSubTypes = 0, Optional ByVal Mandatory As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO, _
                              Optional ByVal defaultvalue As String = "", Optional ByVal Yesno As Boolean = False, Optional ByVal Validvalues() As String = Nothing)
            Dim oUserFieldMD1 As SAPbobsCOM.UserFieldsMD
            oUserFieldMD1 = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
            Try
                'oUserFieldMD1 = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                'If Not (strTab = "OPDN" Or strTab = "OQUT" Or strTab = "OADM" Or strTab = "OPOR" Or strTab = "OWST" Or strTab = "OUSR" Or strTab = "OSRN" Or strTab = "OSPP" Or strTab = "WTR1" Or strTab = "OEDG" Or strTab = "OHEM" Or strTab = "OLCT" Or strTab = "ITM1" Or strTab = "OCRD" Or strTab = "SPP1" Or strTab = "SPP2" Or strTab = "RDR1" Or strTab = "ORDR" Or strTab = "OWHS" Or strTab = "OITM" Or strTab = "INV1" Or strTab = "OWTR" Or strTab = "OWDD" Or strTab = "OWOR" Or strTab = "OWTQ" Or strTab = "OMRV" Or strTab = "JDT1" Or strTab = "OIGN" Or strTab = "OCQG") Then
                '    strTab = "@" + strTab
                'End If
                If Not IsColumnExists(strTab, strCol) Then
                    'If Not oUserFieldMD1 Is Nothing Then
                    '    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD1)
                    'End If
                    'oUserFieldMD1 = Nothing
                    'oUserFieldMD1 = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                    oUserFieldMD1.Description = strDesc
                    oUserFieldMD1.Name = strCol
                    oUserFieldMD1.Type = nType
                    oUserFieldMD1.SubType = nSubType
                    oUserFieldMD1.TableName = strTab
                    oUserFieldMD1.EditSize = nEditSize
                    oUserFieldMD1.Mandatory = Mandatory
                    oUserFieldMD1.DefaultValue = defaultvalue

                    If Yesno = True Then
                        oUserFieldMD1.ValidValues.Value = "Y"
                        oUserFieldMD1.ValidValues.Description = "Yes"
                        oUserFieldMD1.ValidValues.Add()
                        oUserFieldMD1.ValidValues.Value = "N"
                        oUserFieldMD1.ValidValues.Description = "No"
                        oUserFieldMD1.ValidValues.Add()
                    End If

                    Dim split_char() As String
                    If Not Validvalues Is Nothing Then
                        If Validvalues.Length > 0 Then
                            For i = 0 To Validvalues.Length - 1
                                If Trim(Validvalues(i)) = "" Then Continue For
                                split_char = Validvalues(i).Split(",")
                                If split_char.Length <> 2 Then Continue For
                                oUserFieldMD1.ValidValues.Value = split_char(0)
                                oUserFieldMD1.ValidValues.Description = split_char(1)
                                oUserFieldMD1.ValidValues.Add()
                            Next
                        End If
                    End If
                    Dim val As Integer
                    val = oUserFieldMD1.Add
                    If val <> 0 Then
                        objaddon.objapplication.SetStatusBarMessage(objaddon.objcompany.GetLastErrorDescription & " " & strTab & " " & strCol, True)
                    End If
                    'System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD1)
                End If
            Catch ex As Exception
                Throw ex
            Finally

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD1)
                oUserFieldMD1 = Nothing
                GC.WaitForPendingFinalizers()
                GC.Collect()
            End Try
        End Sub

        Private Function IsColumnExists(ByVal Table As String, ByVal Column As String) As Boolean
            Dim oRecordSet As SAPbobsCOM.Recordset
            Dim strSQL As String
            Try
                If objaddon.HANA Then
                    strSQL = "SELECT COUNT(*) FROM CUFD WHERE ""TableID"" = '" & Table & "' AND ""AliasID"" = '" & Column & "'"
                Else
                    strSQL = "SELECT COUNT(*) FROM CUFD WHERE TableID = '" & Table & "' AND AliasID = '" & Column & "'"
                End If

                oRecordSet = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecordSet.DoQuery(strSQL)

                If oRecordSet.Fields.Item(0).Value = 0 Then
                    Return False
                Else
                    Return True
                End If
            Catch ex As Exception
                Throw ex
            Finally
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
                oRecordSet = Nothing
                GC.WaitForPendingFinalizers()
                GC.Collect()
            End Try
        End Function

        Private Sub AddKey(ByVal strTab As String, ByVal strColumn As String, ByVal strKey As String, ByVal i As Integer)
            Dim oUserKeysMD As SAPbobsCOM.UserKeysMD

            Try
                '// The meta-data object must be initialized with a
                '// regular UserKeys object
                oUserKeysMD = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserKeys)

                If Not oUserKeysMD.GetByKey("@" & strTab, i) Then

                    '// Set the table name and the key name
                    oUserKeysMD.TableName = strTab
                    oUserKeysMD.KeyName = strKey

                    '// Set the column's alias
                    oUserKeysMD.Elements.ColumnAlias = strColumn
                    oUserKeysMD.Elements.Add()
                    oUserKeysMD.Elements.ColumnAlias = "RentFac"

                    '// Determine whether the key is unique or not
                    oUserKeysMD.Unique = SAPbobsCOM.BoYesNoEnum.tYES

                    '// Add the key
                    If oUserKeysMD.Add <> 0 Then
                        Throw New Exception(objaddon.objcompany.GetLastErrorDescription)
                    End If
                End If
            Catch ex As Exception
                Throw ex
            Finally
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserKeysMD)
                oUserKeysMD = Nothing
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        End Sub

        Private Sub AddUDO(ByVal strUDO As String, ByVal strUDODesc As String, ByVal nObjectType As SAPbobsCOM.BoUDOObjType, ByVal strTable As String, ByVal childTable() As String, ByVal sFind() As String, _
                           Optional ByVal canlog As Boolean = False, Optional ByVal Manageseries As Boolean = False)

            Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD
            Dim tablecount As Integer = 0
            Try
                oUserObjectMD = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
                If oUserObjectMD.GetByKey(strUDO) = 0 Then

                    oUserObjectMD.Code = strUDO
                    oUserObjectMD.Name = strUDODesc
                    oUserObjectMD.ObjectType = nObjectType
                    oUserObjectMD.TableName = strTable

                    oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO : oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tNO
                    oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO : oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES

                    If Manageseries Then oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES Else oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO

                    If canlog Then
                        oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tYES
                        oUserObjectMD.LogTableName = "A" + strTable.ToString
                    Else
                        oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
                        oUserObjectMD.LogTableName = ""
                    End If

                    oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO : oUserObjectMD.ExtensionName = ""

                    oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
                    tablecount = 1
                    If sFind.Length > 0 Then
                        For i = 0 To sFind.Length - 1
                            If Trim(sFind(i)) = "" Then Continue For
                            oUserObjectMD.FindColumns.ColumnAlias = sFind(i)
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.SetCurrentLine(tablecount)
                            tablecount = tablecount + 1
                        Next
                    End If

                    tablecount = 0
                    If Not childTable Is Nothing Then
                        If childTable.Length > 0 Then
                            For i = 0 To childTable.Length - 1
                                If Trim(childTable(i)) = "" Then Continue For
                                oUserObjectMD.ChildTables.SetCurrentLine(tablecount)
                                oUserObjectMD.ChildTables.TableName = childTable(i)
                                oUserObjectMD.ChildTables.Add()
                                tablecount = tablecount + 1
                            Next
                        End If
                    End If

                    If oUserObjectMD.Add() <> 0 Then
                        Throw New Exception(objaddon.objcompany.GetLastErrorDescription)
                    End If
                End If

            Catch ex As Exception
                Throw ex
            Finally
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD)
                oUserObjectMD = Nothing
                GC.WaitForPendingFinalizers()
                GC.Collect()
            End Try

        End Sub

#End Region

    End Class
End Namespace
