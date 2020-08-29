import pyodbc
import pandas as pd
conn = pyodbc.connect("Driver={SQL Server};"
                              "Server=ussltcsnw1940.solutions.glbsnet.com;"
                              "Database=DATAFIRST;"
                              "UID=deesaw;"
                              "PWD=Welcome@123")
list=[
#,'tvMARA_Relevancy_All_ManufacturerParts_MFRNR_Missing_Rpt'
#,'tvMARA_BroughtMaterials_MAST_WithBOM_Rpt'
#,'tvMARA_Relevancy_1ProductHierarchy_Multiple_Finished_Goods'
#,'tvMARA_EKPO_EKKO_EBELN_Rpt'
#,'tvMAPL_PLKO_PLPO_N_Routing_AUFAK_ScrapMissing_Rpt'
#,'tvMARA_NoBuyMaterials_WithPurchasingView_Rpt'
#,'tvMARA_Sell_Material_ZOldDays_MM_141'
#,'tvMARA_Relevancy_Open_Sales_Order'
#,'tvMARA_Relevancy_SalesOrder_Within_ZOldDays'
#,'tvMARA_VRKME_SalesUoM_Blank_LVSMW_WHUoM_And_AUSMW_UnitOfIssue_NotBlank_Rpt'
#,'tvMARA_NonStockMaterial_NotStocked_With_StorageLoc_Rpt'
'tvMARA_Material_That_Are_Not_Stocked_In_ZOldDays_Rpt'
,'tvMARC_Buy_Material_EKKO_BSTYP_K_Contract_Missing_Rpt'
,'tv_MARC_BESKZ_AUSSS_INHOUSE_NO_ASSEMBLYSCRAP_Rpt'
,'tvMARA_Stock_Material_PSTAT_L_HasStorageView_LGNUM_Blank_NoWarehouse_Rpt'
     ]

for v in list:
    print(v)
    script="select * from dbo."+v+";"
    cursor = conn.cursor()
    script="select * from dbo."+v 
    cursor.execute(script)
    rows=cursor.fetchall() 
    names = [desc[0] for desc in cursor.description] 
    df = pd.DataFrame([tuple(t) for t in rows]) 
    df.columns=names
    print(df.shape)
    excelname=v+'.xlsx'
    writer = pd.ExcelWriter(excelname)
    df.to_excel(writer, sheet_name='bar',index=False)
    writer.save()