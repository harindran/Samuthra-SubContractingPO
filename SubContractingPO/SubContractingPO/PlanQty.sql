-----------------------Sub-Contracting Add-on SP-------------
IF ((:transaction_type=(n'A') OR :transaction_type=(n'U')) AND :object_type='SUBPO') THEN
Select (Select distinct 1  From "@MIPL_OPOR" T0 join "@MIPL_POR2" T1 on T0."DocEntry"=T1."DocEntry" 
Where (ifnull(T1."U_DocDate",'')='' or ifnull(T1."U_RefNo",'')='') and T0."DocEntry"=:list_of_cols_val_tab_del )
INTO temp_var_0 FROM DUMMY;
  	IF :temp_var_0 > 0  THEN
 	  error := 10001;
	  error_message := 'Document Date and Reference Number is mandatory in Output Tab';
	End If;
END IF;

IF (:transaction_type=(n'U') AND :object_type='SUBPO')  Then
Select (Select distinct 1 From "@MIPL_OPOR" T0 join "@MIPL_POR3" T1 on T0."DocEntry"=T1."DocEntry" 
Where (ifnull(T1."U_DocDate",'')='' or ifnull(T1."U_RefNo",'')='') and T0."DocEntry"=:list_of_cols_val_tab_del )
INTO temp_var_1 FROM DUMMY;
  		IF :temp_var_1 > 0  Then
    	error := 10002;
		error_message := 'Document Date and Reference Number is mandatory in Scrap Tab';
		END IF;
END IF;
-------------------------------
-------------------------------Batch Query-----------------------------------------------
Alter PROCEDURE "MIPL_GetBatch"(IN SubEntry VARCHAR(100),In DocEntry VARCHAR(1000),In ItemCode VARCHAR(30),In WhsCode VARCHAR(30)) 
LANGUAGE SQLSCRIPT SQL SECURITY INVOKER
 AS
BEGIN 
  Select A."U_SubConNo",A."BatchSerial",A."WhsCode",A."DocEntry",A."ItemCode",Sum(A."Quantity") as "Qty",A."Status" from (
SELECT distinct T0."U_SubConNo",I1."BatchNum" "BatchSerial",T4."WhsCode",T1."DocEntry",T1."ItemCode",T4."Quantity" as "Qty", 
I1."Quantity",T4."Status" from OWTR T0 left join WTR1 T1 on T0."DocEntry"=T1."DocEntry"
left outer join IBT1 I1 on T1."ItemCode"=I1."ItemCode"   and (T1."DocEntry"=I1."BaseEntry" and T1."ObjType"=I1."BaseType") 
and T1."LineNum"=I1."BaseLinNum"
left outer join OIBT T4 on T4."ItemCode"=I1."ItemCode" and I1."BatchNum"=T4."BatchNum" and I1."WhsCode" = T4."WhsCode"
where I1."Direction"=0
union all
SELECT distinct T0."U_SubConNo",I1."BatchNum" "BatchSerial",T4."WhsCode",T1."DocEntry",T1."ItemCode",T4."Quantity" as "Qty", 
-I1."Quantity",T4."Status" from OWTR T0 left join WTR1 T1 on T0."DocEntry"=T1."DocEntry"
left outer join IBT1 I1 on T1."ItemCode"=I1."ItemCode"   and (T1."DocEntry"=I1."BaseEntry" and T1."ObjType"=I1."BaseType") 
and T1."LineNum"=I1."BaseLinNum"
left outer join OIBT T4 on T4."ItemCode"=I1."ItemCode" and I1."BatchNum"=T4."BatchNum" and I1."WhsCode" = T4."WhsCode"
where I1."Direction"=1)A 
Where A."U_SubConNo"=:SubEntry and A."DocEntry" in (:DocEntry) and A."ItemCode"=:ItemCode and A."WhsCode"=:WhsCode 
and A."BatchSerial" <>'' and A."Status"=0 Group by  A."U_SubConNo",A."BatchSerial",A."WhsCode",A."DocEntry",A."ItemCode",A."Status" ;
END;
--------------FMS for Purchase order Item Selection ----------------
----------------------------------------
select T1."LineNum",T1."ItemCode",T1."Dscription",T1."Quantity",T1."Price",T1."LineTotal",T1."WhsCode" from OPOR T0 join POR1 T1 on T0."DocEntry"=T1."DocEntry" 
where T1."DocEntry"=$[$SubPoNum.0.0] and T1."LineStatus"='O'
-------------Updated SP Jan 13 2021

Create PROCEDURE "MIPL_ValidateGRQty"(IN Line VARCHAR(100),In DocEntry VARCHAR(10),In ItemCode VARCHAR(30)) 
LANGUAGE SQLSCRIPT SQL SECURITY INVOKER
 AS
BEGIN 
DECLARE T VARCHAR(5000);      
 Select 'Select  Distinct T1."U_SQty"-T4."Qty" as "PendQty"from  "@MIPL_OPOR" T1 inner join "'||:Line||'" T0 on T0."DocEntry"=T1."DocEntry" and T0."U_Itemcode"=T1."U_SItemCode" inner  join
(select T2."U_SubConNo",sum(T3."Quantity") as "Qty",T3."ItemCode" from  OIGN T2 join IGN1 T3 on T2."DocEntry"=T3."DocEntry" group by T2."U_SubConNo",T3."ItemCode" ) as T4
 on T4."U_SubConNo"=T1."DocEntry" and T0."U_Itemcode"=T4."ItemCode" where T1."U_SItemCode"='''||:ItemCode||'''  and T1."DocEntry"='||:DocEntry||''
 
  INTO     T 
  FROM    DUMMY; 
  EXECUTE IMMEDIATE :T; 
END;


Create PROCEDURE "MIPL_GetPendQty_Inv_GI_GR"( IN Header  VARCHAR(100), 
                                      IN Line VARCHAR(100),In DocEntry Integer) LANGUAGE SQLSCRIPT SQL SECURITY INVOKER AS
BEGIN 
  DECLARE T VARCHAR(5000);
  
 Select ' select Case when Sum(T2."Quantity") >0 then T0."U_PlanQty"-Sum(T2."Quantity") else T0."U_PlanQty" end as "PendQty",T0."U_Itemcode"
from "@MIPL_POR1" T0  left join "@MIPL_OPOR" T1 on T0."DocEntry"=T1."DocEntry" 
left join "'||:Header||'" T3 on T3."U_SubConNo"=T1."DocEntry" left join "'||:Line||'" T2  on T3."DocEntry"=T2."DocEntry" and T0."U_Itemcode"=T2."ItemCode"
where T1."DocEntry"='||:DocEntry||' group by T0."U_PlanQty",T0."U_Itemcode" order by "PendQty" '
  INTO     T 
  FROM    DUMMY;
EXECUTE IMMEDIATE :T; 
END;

---------------------------------------------




--------------FMS for JEGL In UDO Screen ----------------
select "AcctName" from OACT where "AcctCode"=$["@MIPL_SBGL"."U_JEGLCode"]

Create PROCEDURE "MIPL_ValidateGRQty"(IN Line VARCHAR(100),In DocNum VARCHAR(15),In ItemCode VARCHAR(30)) 
LANGUAGE SQLSCRIPT SQL SECURITY INVOKER
 AS
BEGIN 
DECLARE T VARCHAR(5000);      
 Select 'Select  Distinct T1."U_SQty"-T4."Qty" as "PendQty"from  "@MIPL_OPOR" T1 inner join "'||:Line||'" T0 on T0."DocEntry"=T1."DocEntry" and T0."U_Itemcode"=T1."U_SItemCode" inner  join
(select T2."U_SubConNo",sum(T3."Quantity") as "Qty",T3."ItemCode" from  OIGN T2 join IGN1 T3 on T2."DocEntry"=T3."DocEntry" group by T2."U_SubConNo",T3."ItemCode" ) as T4
 on T4."U_SubConNo"=T1."DocEntry" and T0."U_Itemcode"=T4."ItemCode" where T1."U_SItemCode"='''||:ItemCode||'''  and T1."DocNum"='||:DocNum||''
 
  INTO     T 
  FROM    DUMMY; 
  EXECUTE IMMEDIATE :T; 
END;

CREATE PROCEDURE "MIPL_GetPendQty_Inv_GI_GR"( IN Header  VARCHAR(100), 
                                      IN Line VARCHAR(100),In DocNum Integer) LANGUAGE SQLSCRIPT SQL SECURITY INVOKER AS
BEGIN 
  DECLARE T VARCHAR(5000);
  
 Select ' select Case when Sum(T2."Quantity") >0 then T0."U_PlanQty"-Sum(T2."Quantity") else T0."U_PlanQty" end as "PendQty",T0."U_Itemcode"
from "@MIPL_POR1" T0  left join "@MIPL_OPOR" T1 on T0."DocEntry"=T1."DocEntry" 
left join "'||:Header||'" T3 on T3."U_SubConNo"=T1."DocEntry" left join "'||:Line||'" T2  on T3."DocEntry"=T2."DocEntry" and T0."U_Itemcode"=T2."ItemCode"
where T1."DocNum"='||:DocNum||' group by T0."U_PlanQty",T0."U_Itemcode" order by "PendQty" '
  INTO     T 
  FROM    DUMMY;
EXECUTE IMMEDIATE :T; 
END;

--------------------------------------------------Latest Stored Procedure Oct 22 2020----------------
Alter PROCEDURE "MIPL_GetPendQty_Inv_GI_GR"( IN Header  VARCHAR(100), 
                                      IN Line VARCHAR(100),In DocNum Integer) LANGUAGE SQLSCRIPT SQL SECURITY INVOKER AS
BEGIN 
  DECLARE T VARCHAR(5000);
  
 Select ' select Case when Sum(T2."Quantity") >0 then T0."U_PlanQty"-Sum(T2."Quantity") else T0."U_PlanQty" end as "PendQty",T0."U_Itemcode"
from "@MIPL_POR1" T0  left join "@MIPL_OPOR" T1 on T0."DocEntry"=T1."DocEntry" 
left join "'||:Header||'" T3 on T3."U_SubConNo"=T1."DocEntry" left join "'||:Line||'" T2  on T3."DocEntry"=T2."DocEntry" and T0."U_Itemcode"=T2."ItemCode"
where T1."DocNum"='||:DocNum||' group by T0."U_PlanQty",T0."U_Itemcode" order by "PendQty" '
  INTO     T 
  FROM    DUMMY;
EXECUTE IMMEDIATE :T; 
END;

Alter PROCEDURE "MIPL_ValidateGRQty"(IN Line VARCHAR(100),In DocNum VARCHAR(10),In ItemCode VARCHAR(30)) 
LANGUAGE SQLSCRIPT SQL SECURITY INVOKER
 AS
BEGIN 
DECLARE T VARCHAR(5000);      
 Select 'Select  Distinct T1."U_SQty"-T4."Qty" as "PendQty"from  "@MIPL_OPOR" T1 inner join "'||:Line||'" T0 on T0."DocEntry"=T1."DocEntry" and T0."U_Itemcode"=T1."U_SItemCode" inner  join
(select T2."U_SubConNo",sum(T3."Quantity") as "Qty",T3."ItemCode" from  OIGN T2 join IGN1 T3 on T2."DocEntry"=T3."DocEntry" group by T2."U_SubConNo",T3."ItemCode" ) as T4
 on T4."U_SubConNo"=T1."DocEntry" and T0."U_Itemcode"=T4."ItemCode" where T1."U_SItemCode"='''||:ItemCode||'''  and T1."DocNum"='||:DocNum||''
 
  INTO     T 
  FROM    DUMMY; 
  EXECUTE IMMEDIATE :T; 
END;
------------------------------------------------
-------------------	-----------Latest Stored Procedure Sep 26 2020---------------------------
CREATE PROCEDURE "MIPL_GetPendQty_Inv_GI_GR"( IN Header  VARCHAR(100), 
                                      IN Line VARCHAR(100),In DocNum Integer) LANGUAGE SQLSCRIPT SQL SECURITY INVOKER AS
BEGIN 
  DECLARE T VARCHAR(5000);
  
 Select ' select Case when Sum(T2."Quantity") >0 then T0."U_PlanQty"-Sum(T2."Quantity") else T0."U_PlanQty" end as "PendQty",T0."U_Itemcode"
from "@MIPL_POR1" T0  left join "@MIPL_OPOR" T1 on T0."DocEntry"=T1."DocEntry" 
left join "'||:Header||'" T3 on T3."U_SubConNo"=T1."DocEntry" left join "'||:Line||'" T2  on T3."DocEntry"=T2."DocEntry" and T0."U_Itemcode"=T2."ItemCode"
where T1."DocNum"='||:DocNum||' group by T0."U_PlanQty",T0."U_Itemcode" order by "PendQty" '
  INTO     T 
  FROM    DUMMY;
EXECUTE IMMEDIATE :T; 
END;


CREATE PROCEDURE "MIPL_ValidateGRQty"(IN Line VARCHAR(100),In DocNum VARCHAR(10),In ItemCode VARCHAR(30)) 
LANGUAGE SQLSCRIPT SQL SECURITY INVOKER
 AS
BEGIN 
DECLARE T VARCHAR(5000);      
 Select 'Select  Distinct T1."U_SQty"-T4."Qty" as "PendQty"from  "@MIPL_OPOR" T1 inner join "'||:Line||'" T0 on T0."DocEntry"=T1."DocEntry" and T0."U_Itemcode"=T1."U_SItemCode" inner  join
(select T2."U_SubConNo",sum(T3."Quantity") as "Qty",T3."ItemCode" from  OIGN T2 join IGN1 T3 on T2."DocEntry"=T3."DocEntry" group by T2."U_SubConNo",T3."ItemCode" ) as T4
 on T4."U_SubConNo"=T1."DocEntry" and T0."U_Itemcode"=T4."ItemCode" where T1."U_SItemCode"='''||:ItemCode||'''  and T1."DocNum"='||:DocNum||''
 
  INTO     T 
  FROM    DUMMY; 
  EXECUTE IMMEDIATE :T; 
END;
------------------------------------------------------------
-------------------	----------- Stored Procedure Sep 11 2020---------------------------
Create PROCEDURE "MIPL_ValidateGRQty"(IN Line VARCHAR(100),In DocNum VARCHAR(15),In ItemCode VARCHAR(30)) 
LANGUAGE SQLSCRIPT SQL SECURITY INVOKER
 AS
BEGIN 
DECLARE T VARCHAR(5000);      
 Select 'Select  Distinct T1."U_SQty"-T4."Qty" as "PendQty"from  "@MIPL_OPOR" T1 inner join "'||:Line||'" T0 on T0."DocEntry"=T1."DocEntry" and T0."U_Itemcode"=T1."U_SItemCode" inner  join
(select T2."U_SubConNo",sum(T3."Quantity") as "Qty",T3."ItemCode" from  OIGN T2 join IGN1 T3 on T2."DocEntry"=T3."DocEntry" group by T2."U_SubConNo",T3."ItemCode" ) as T4
 on T4."U_SubConNo"=T1."DocEntry" and T0."U_Itemcode"=T4."ItemCode" where T1."U_SItemCode"='''||:ItemCode||'''  and T1."DocNum"='||:DocNum||''
 
  INTO     T 
  FROM    DUMMY; 
  EXECUTE IMMEDIATE :T; 
END;

Alter PROCEDURE "MIPL_GetPendQty_Inv_GI_GR"( IN Header  VARCHAR(100), 
                                      IN Line VARCHAR(100),In DocNum Integer) LANGUAGE SQLSCRIPT SQL SECURITY INVOKER AS
BEGIN 
  DECLARE T VARCHAR(5000);
  
 Select ' select Case when Sum(T2."Quantity") >0 then T0."U_PlanQty"-Sum(T2."Quantity") else T0."U_PlanQty" end as "PendQty",T0."U_Itemcode"
from "@MIPL_POR1" T0  left join "@MIPL_OPOR" T1 on T0."DocEntry"=T1."DocEntry" 
left join "'||:Header||'" T3 on T3."U_SubConNo"=T1."DocEntry" left join "'||:Line||'" T2  on T3."DocEntry"=T2."DocEntry" and T0."U_Itemcode"=T2."ItemCode"
where T1."DocNum"='||:DocNum||' group by T0."U_PlanQty",T0."U_Itemcode" order by "PendQty" '
  INTO     T 
  FROM    DUMMY;
EXECUTE IMMEDIATE :T; 
END;
--------------------------------------------------------
Create PROCEDURE "MIPL_ValidateGRQty"(IN Line VARCHAR(100),In DocNum VARCHAR(10),In ItemCode VARCHAR(30)) 
LANGUAGE SQLSCRIPT SQL SECURITY INVOKER
 AS
BEGIN 
DECLARE T VARCHAR(5000);      
 Select 'Select  Distinct T1."U_SQty"-T4."Qty" as "PendQty"from  "@MIPL_OPOR" T1 inner join "'||:Line||'" T0 on T0."DocEntry"=T1."DocEntry" and T0."U_Itemcode"=T1."U_SItemCode" inner  join
(select T2."U_SubConNo",sum(T3."Quantity") as "Qty",T3."ItemCode" from  OIGN T2 join IGN1 T3 on T2."DocEntry"=T3."DocEntry" group by T2."U_SubConNo",T3."ItemCode" ) as T4
 on T4."U_SubConNo"=T1."DocNum" and T0."U_Itemcode"=T4."ItemCode" where T1."U_SItemCode"='''||:ItemCode||'''  and T1."DocNum"='||:DocNum||''
 
  INTO     T 
  FROM    DUMMY; 
  EXECUTE IMMEDIATE :T; 
END;
------------------updated-------------------------

CREATE PROCEDURE "MIPL_GetPendQty_Inv_GI_GR"( IN Header  VARCHAR(100), 
                                      IN Line VARCHAR(100),In DocNum Integer) LANGUAGE SQLSCRIPT SQL SECURITY INVOKER AS
BEGIN 
  DECLARE T VARCHAR(5000);
  
 Select ' select Case when Sum(T2."Quantity") >0 then T0."U_PlanQty"-Sum(T2."Quantity") else T0."U_PlanQty" end as "PendQty",T0."U_Itemcode"
from "@MIPL_POR1" T0  left join "@MIPL_OPOR" T1 on T0."DocEntry"=T1."DocEntry" 
left join "'||:Header||'" T3 on T3."U_SubConNo"=T1."DocNum" left join "'||:Line||'" T2  on T3."DocEntry"=T2."DocEntry" and T0."U_Itemcode"=T2."ItemCode"
where T1."DocNum"='||:DocNum||' group by T0."U_PlanQty",T0."U_Itemcode" order by "PendQty"'
  INTO     T 
  FROM    DUMMY;
EXECUTE IMMEDIATE :T; 
END;

call "MIPL_GetPendQty_Inv_GI_GR"('OWTR','WTR1','53')


 Create PROCEDURE "MIPL_ValidateGRQty"(IN Line VARCHAR(100),In DocNum VARCHAR(10),In ItemCode VARCHAR(10)) 
LANGUAGE SQLSCRIPT SQL SECURITY INVOKER
 AS
BEGIN 
DECLARE T VARCHAR(5000);
       Select 'Select  1 as "Status"
 from  "@MIPL_OPOR" T1 join "'||:Line||'" T0 on T0."DocEntry"=T1."DocEntry" 
 left join OIGN T3 on T3."U_SubConNo"=T1."DocNum"  join IGN1 T2  on T3."DocEntry"=T2."DocEntry" and T0."U_Itemcode"=T2."ItemCode"
 where T3."U_SubConNo"='||:DocNum||'  and T2."ItemCode"='''||:ItemCode||''' and  ifnull(T0."U_GRNo",'''')<>'''' and T0."U_Status"=''C''
 group by T1."U_SQty" having T1."U_SQty"-Sum(T2."Quantity")<=0'
  INTO     T 
  FROM    DUMMY; 
  EXECUTE IMMEDIATE :T; 
END;






------------------------------------Old-------------------------------------
CREATE PROCEDURE "MIPL_GetPendingQty"( IN Line VARCHAR(100),In DocNum VARCHAR(10),IN TabType VARCHAR(10))
LANGUAGE SQLSCRIPT SQL SECURITY INVOKER 
AS
BEGIN 
  DECLARE T VARCHAR(5000);
  Select 'select Case when Sum(T2."Quantity") >0 then T0."U_Qty"-Sum(T2."Quantity") else T0."U_Qty" end as "PendQty",T0."U_Itemcode",
  T0."U_Price",T0."U_LineTot",T0."U_WhsCode"
 from "'||:Line||'" T0  left join "@MIPL_OPOR" T1 on T0."DocEntry"=T1."DocEntry" 
 left join OIGN T3 on T3."U_SubConNo"=T1."DocNum" left join IGN1 T2  on T3."DocEntry"=T2."DocEntry" and T0."U_Itemcode"=T2."ItemCode"
 where T1."DocNum"='||:DocNum||' and T0."U_TabType"='''||:TabType||'''  group by T0."U_Qty",T0."U_Itemcode",T0."U_Price",T0."U_LineTot",T0."U_WhsCode"   '
  INTO     T 
  FROM    DUMMY; 
  EXECUTE IMMEDIATE :T; 
END;

Create PROCEDURE "MIPL_ValidateGRQty"(IN Line VARCHAR(100),In DocNum VARCHAR(10),In ItemCode VARCHAR(10)) 
LANGUAGE SQLSCRIPT SQL SECURITY INVOKER
 AS
BEGIN 
DECLARE T VARCHAR(5000);
   Select 'Select 1 as "Status"
 from "'||:Line||'" T0  left join "@MIPL_OPOR" T1 on T0."DocEntry"=T1."DocEntry" 
 left join OIGN T3 on T3."U_SubConNo"=T1."DocNum" left join IGN1 T2  on T3."DocEntry"=T2."DocEntry" and T0."U_Itemcode"=T2."ItemCode"
 where T1."DocNum"='||:DocNum||'  and T0."U_Itemcode"='''||:ItemCode||''' 
 group by T1."U_SQty" having T1."U_SQty"-Sum(T2."Quantity")<=0 '
  INTO     T 
  FROM    DUMMY; 
  EXECUTE IMMEDIATE :T; 
END;

Alter PROCEDURE "MIPL_ValidatePlannedQty"( IN Header  VARCHAR(100), 
                                      IN Line VARCHAR(100),In DocNum Integer) LANGUAGE SQLSCRIPT SQL SECURITY INVOKER AS
BEGIN 
  DECLARE T VARCHAR(5000);
 Select 'select distinct 1 as "Status",T2."ItemCode"
from "@MIPL_POR1" T0  left join "@MIPL_OPOR" T1 on T0."DocEntry"=T1."DocEntry" 
 left join "'||:Header||'" T3 on T3."U_SubConNo"=T1."DocNum" 
  left join "'||:Line||'" T2  on T3."DocEntry"=T2."DocEntry" and T0."U_Itemcode"=T2."ItemCode" and T0."U_SubWhse"=T2."WhsCode"
where T1."DocNum"='||:DocNum||' group by T0."U_PlanQty", T2."ItemCode" having T0."U_PlanQty"-Sum(T2."Quantity")<=0'
  INTO     T 
  FROM    DUMMY;
EXECUTE IMMEDIATE :T; 
END;

Call "MIPL_ValidatePlannedQty" ('OWTR','WTR1',20)
Call "MIPL_GetPendingQty" ('@MIPL_POR2','25','Output')


---------------------------------------------------------------------------------Rough Query ----------------------------------------------------


select Case when Sum(T2."Quantity") >0 then T0."U_PlanQty"-Sum(T2."Quantity") else T0."U_PlanQty" end as "PendQty",T0."U_Itemcode",
T0."U_WhsCode",T0."U_SubWhse"
from "@MIPL_POR1" T0  left join "@MIPL_OPOR" T1 on T0."DocEntry"=T1."DocEntry" 
 left join owtr T3 on T3."U_SubConNo"=T1."DocNum" 
  left join wtr1 T2  on T3."DocEntry"=T2."DocEntry" and T0."U_Itemcode"=T2."ItemCode"
where T1."DocNum"=26  group by T0."U_PlanQty",T0."U_Itemcode",T0."U_WhsCode",T0."U_SubWhse"

select Case when Sum(T2."Quantity") >0 then T0."U_Qty"-Sum(T2."Quantity") else T0."U_Qty" end as "PendQty",T0."U_Itemcode"
 from "@MIPL_POR2" T0  left join "@MIPL_OPOR" T1 on T0."DocEntry"=T1."DocEntry" 
 left join OIGN T3 on T3."U_SubConNo"=T1."DocNum" left join IGN1 T2  on T3."DocEntry"=T2."DocEntry" and T0."U_Itemcode"=T2."ItemCode"
 where T1."DocNum"=25 and  T0."U_Itemcode"<>'' and T2."U_TabType"='Output'  group by T0."U_Qty",T0."U_Itemcode"  -- and T2."U_TabType"='Output'
 
 select Case when Sum(T2."Quantity") >0 then T0."U_Qty"-Sum(T2."Quantity") else T0."U_Qty" end as "PendQty",T0."U_Itemcode"
 from "@MIPL_POR3" T0  left join "@MIPL_OPOR" T1 on T0."DocEntry"=T1."DocEntry" 
 left join OIGN T3 on T3."U_SubConNo"=T1."DocNum" left join IGN1 T2  on T3."DocEntry"=T2."DocEntry" and T0."U_Itemcode"=T2."ItemCode"
 where T1."DocNum"=25 and  T0."U_Itemcode"<>'' and T2."U_TabType"='Scrap'  group by T0."U_Qty",T0."U_Itemcode"
 
 select T1."U_SQty"-ifnull(Sum(T3."Quantity"),0) "OpenQty"  from "@MIPL_POR2" T0 left join "@MIPL_OPOR" T1 on T0."DocEntry"=T1."DocEntry"
left join OIGN T2 on T1."DocNum"=T2."U_SubConNo" left join IGN1 T3 on T2."DocEntry"=T3."DocEntry" and  T0."U_Itemcode"=T3."ItemCode"  and T1."U_SItemCode" =T3."ItemCode"
where T1."U_SItemCode"='STFB003'  and T1."DocNum"=35  group by T1."U_SQty" 

select T1."U_SQty"-ifnull(Sum(T3."Quantity"),0) "VOBStock"  from "@MIPL_POR2" T0 left join "@MIPL_OPOR" T1 on T0."DocEntry"=T1."DocEntry"
left join OIGN T2 on T1."DocNum"=T2."U_SubConNo" left join IGN1 T3 on T2."DocEntry"=T3."DocEntry" and  T0."U_Itemcode"=T3."ItemCode"  and T1."U_SItemCode" =T3."ItemCode"
where T1."U_SItemCode"='STFB003' and T1."U_CardCode"='V0007'   group by T1."U_SQty" 


Select Distinct 1 from "@MIPL_POR2" T1 left join  "@MIPL_OPOR" T0 on T0."DocEntry"=T1."DocEntry" 
 left join OIGN T2 on T0."DocNum"=T2."U_SubConNo" left join IGN1 T3 on T2."DocEntry"=T3."DocEntry"   and  T1."U_Itemcode"=T3."ItemCode"  and T0."U_SItemCode" =T3."ItemCode"
    where T0."U_SItemCode"='STFB003' and T0."Status"='O' and T0."DocNum"='36' 
    group by T0."U_SQty" having T0."U_SQty"- sum(T3."Quantity")<=0;  
 ----------------------------------------------
 Select * from (
SELECT distinct T0."U_SubConNo",I1."BatchNum" "BatchSerial",T1."DocEntry",T1."ItemCode", I1."Quantity" from OWTR T0 left join WTR1 T1 on T0."DocEntry"=T1."DocEntry"
left outer join IBT1 I1 on T1."ItemCode"=I1."ItemCode"   and (T1."DocEntry"=I1."BaseEntry" and T1."ObjType"=I1."BaseType") and T1."LineNum"=I1."BaseLinNum"
left outer join OIBT T4 on T4."ItemCode"=I1."ItemCode" and I1."BatchNum"=T4."BatchNum" and I1."WhsCode" = T4."WhsCode"
)A Where A."U_SubConNo"='35' and A."DocEntry"='87' and A."ItemCode"='SLMA003' and A."BatchSerial" <>''
 ----------------------------------------------
