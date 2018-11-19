using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace ExcelDataImporter
{
    public static class ImportFromExcel
    {
        public static DataTable ImportDataFromExcelByOleDb(bool hasTitle=false)
        {
            OpenFileDialog fopen = new OpenFileDialog();
            fopen.Filter = "Excel(*.xlsx)|*.xlsx|Excel(*.xls)|*.xls";
            fopen.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            fopen.Multiselect = false;


            if (fopen.ShowDialog() == false)
                return null;

            var filePath = fopen.FileName;

            var fileType = Path.GetExtension(filePath);

            if (string.IsNullOrEmpty(fileType))
                return null;

            using (DataSet ds=new DataSet())
            {
                // support Excel 2003(Excel 8.0) / Excel 2007 and above(Excel 12.0)
                string strCon = string.Format("Provider=Microsoft.Jet.OLEDB.{0}.0;"+"Extended Properties=\"Excel {1}.0;HDR={2};IMEX=1;\";"+"data source={3};",
                                                            (fileType==".xls"?4:12),(fileType==".xls"?8:12),(hasTitle?"Yes":"No"),filePath);
                string strCom = " SELECT * FROM [Sheet1$]";

                using (OleDbConnection myConn = new OleDbConnection(strCon))
                    using(OleDbDataAdapter myCommand=new OleDbDataAdapter(strCom, myConn))
                {
                    myConn.Open();
                    myCommand.Fill(ds);
                }
                if (ds == null || ds.Tables.Count <= 0)
                    return null;
                return ds.Tables[0];
            }

        }
    }
}
