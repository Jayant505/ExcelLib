using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelManageLib.DataModel
{
    public static class ToolConfiguration
    {
       public static String areaTemplateExcelFilePath = @"D:\Marco\Template\Area_Template.xlsx";

       public static String ecAdditionExcelFilePath = @"D:\Marco\四区货号全新\EC_addition.xlsx";
       public static String ncAdditionExcelFilePath = @"D:\Marco\四区货号全新\NC_addition.xlsx";
       public static String scAdditionExcelFilePath = @"D:\Marco\四区货号全新\SC_addition.xlsx";
       public static String wcAdditionExcelFilePath = @"D:\Marco\四区货号全新\WC_addition.xlsx";

       public static String ecExtensionExcelFilePath = @"D:\Marco\四区货号级别上新\EC_extension.xlsx";
       public static String ncExtensionExcelFilePath = @"D:\Marco\四区货号级别上新\NC_extension.xlsx";
       public static String scExtensionExcelFilePath = @"D:\Marco\四区货号级别上新\SC_extension.xlsx";
       public static String wcExtensionExcelFilePath = @"D:\Marco\四区货号级别上新\WC_extension.xlsx";

       public static String additionAndExtensionExcelPath = @"D:\Marco\Additional&Extention-2019(重要勿删).xlsx";
       public static String additionSheetName = "2018 addition";
       public static String extendSheetName = "2018 Store Extension";

       public static String storeListExcelFilePath = @"D:\Marco\Store_List-全国.xlsx";
       public static String storeListSheetName = "南区&东区店铺信息";
       
       public static String summaryExcelFilePath = @"D:\Marco\爆图汇总_20190109_141804.xlsx";
       public static String futureItemAdditionSheetName = "爆图汇总_future_item_Addition";
       public static String futureItemExtensionSheetName = "爆图汇总_future_item_Extension";
       public static String itemExtensionSheetName = "爆图汇总_item_Extension";

        public static String weekSheetName = "Sheet1";
        public static String configurationExcelFilePath = @"D:\Marco\新增图流程_参数配置表Test.xlsx";

        //public static String ordingExcelFilePath = @"D:\Marco\OrderingTest.xlsx"; 
        
        public static String ordingExcelFilePath = @"D:\Marco\Ordering-update.xlsb";
        //public static String ordingExcelFilePath = @"D:\Marco\Ordering- for m.xlsb";
        public static String orderingExcelPassword = "5201314";


        public static String ordingSheetName = "Sheet1";

       public static String LogPath = @"C:\TestExcelToolLog\";
       public static bool isWriteLog = true;

       // ////========================
       // public static String testFile = @"D:\Marco\Test.xlsx";





    }
}
