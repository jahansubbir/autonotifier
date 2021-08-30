using System.Data.OleDb;
using System.IO;

namespace Autonotifier2.DataAccess
{
    public class Gateway
    {
        private string connectionString = "";
        public string Query { get; set; }
        public OleDbConnection Connection { get; set; }
        public OleDbCommand Command { get; set; }
        public OleDbDataReader Reader { get; set; }


        public Gateway(/*string fileName, string extension*/)
        {
           

        }

        protected void Connect(string fileName)
        {
            if (Path.GetExtension(fileName).CompareTo(".xls") == 0)
            {
                connectionString = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + fileName + "';Extended Properties='Excel 8.0;HDR=Yes;';"; //for below excel 2007  
            }
            
            else 
            {
                connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + fileName + "';Extended Properties='Excel 12.0;HDR=YES';"; //for above excel 2007

            }
            Connection = new OleDbConnection(connectionString);

        }
        /*public DataTable ReadExcel(string fileName, string extension)
        {
            if (extension.CompareTo(".xls") == 0)
                connectionString = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007  
            else
                connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=NO';"; //for above excel 2007  
            using (OleDbConnection con = new OleDbConnection(conn))
            {
                try
                {
                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter("select * from [Sheet1$]", con); //here we read data from sheet1  
                    oleAdpt.Fill(dtexcel); //fill excel data into dataTable  
                }
                catch { }
            }
            return dtexcel;
        }*/
    }
}
