// See https://aka.ms/new-console-template for more information
using System.IO;
namespace TableTransfer
{
    class Nonsense
    {
        static void Main(string[] args)
        {
            using (StreamReader sr = new StreamReader("mode.cfg"))
            {
                TableTransfer tt = new TableTransfer("db.cfg", "IP.xlsx");
                if (sr.ReadLine() != null)
                {
                    switch (sr.ReadLine())
                    {
                        case "0":
                            tt.MySQLToExcel();
                            break;
                        case "1":
                            tt.ExcelToMySQL();
                            break;
                        default:
                            break;
                    }
                }
            }
        }
    }
}
