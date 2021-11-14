using System;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace EBSExcelOkumaClass
{
    public class ExcellRead
    {
        private static readonly OleDbConnection oleDbConnection = new OleDbConnection();
        public OleDbConnection EBSBaglanti = oleDbConnection;
        private OleDbDataAdapter EBSAdtr;
        private DataTable dt;

        public static string yol = "";
        public OleDbConnection Conneciton()
        {
            EBSBaglanti.ConnectionString = $"Provider = Microsoft.ACE.OLEDB.12.0; Data Source ='{yol}'; Extended Properties = Excel 12.0;";

            if (EBSBaglanti.State == ConnectionState.Closed)
            {
                EBSBaglanti.Open();
            }
      
            return EBSBaglanti;
        }

        public  void Excelverioku(string tsql,DataGridView EBSdtgrd)
        {
            dt = new DataTable();
            EBSAdtr = new OleDbDataAdapter(tsql,Conneciton());
            EBSAdtr.Fill(dt);
            EBSdtgrd.DataSource = dt;
        }

        public static void GetEBSSayfaAdiAl(ComboBox combo)
        {
            Microsoft.Office.Interop.Excel.Application EBSxlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook EBSExcelSayfa = EBSxlApp.Workbooks.Open(yol);

            string[] excelSheets = new String[EBSExcelSayfa.Worksheets.Count];
            int i = 0;
            foreach (Microsoft.Office.Interop.Excel.Worksheet wSheet in EBSExcelSayfa.Worksheets)
            {
                excelSheets[i] = wSheet.Name;
                i++;
            }
            foreach (string  EBSSayfaAdi in excelSheets)
            {
                combo.Items.Add(EBSSayfaAdi);
            }
        }
    }
  }

