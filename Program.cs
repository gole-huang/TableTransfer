// See https://aka.ms/new-console-template for more information
namespace TableTransfer
{
    class Nonsense
    {
        static void Main(string[] args)
        {
            TableTransfer tt = new TableTransfer("db.cfg","IP.xlsx");
            tt.MySQLToExcel();
        }
    }
}
