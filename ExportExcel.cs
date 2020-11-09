using System.Data.SqlClient;
using System.IO;
using ClosedXML.Excel;

internal static partial class ExportExcel
{
    public static void ExceleAktar(string FileName, string SQL)
    {
        using (var con = new SqlConnection(ConnectionString))
        {
            using (var cmd = new SqlCommand(SQL))
            {
                using (var sda = new SqlDataAdapter())
                {
                    cmd.Connection = con;
                    sda.SelectCommand = cmd;
                    using (var dt = new DataTable())
                    {
                        sda.Fill(dt);
                        using (var wb = new XLWorkbook())
                        {
                            wb.Worksheets.Add(dt, "SheetName1");
                            HttpContext.Current.Response.Clear();
                            HttpContext.Current.Response.Buffer = true;
                            HttpContext.Current.Response.Charset = "";
                            HttpContext.Current.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                            HttpContext.Current.Response.AddHeader("content-disposition", "attachment;filename=" + FileName + ".xlsx");
                            using (var MyMemoryStream = new MemoryStream())
                            {
                                wb.SaveAs(MyMemoryStream);
                                MyMemoryStream.WriteTo(HttpContext.Current.Response.OutputStream);
                                HttpContext.Current.Response.Flush();
                                HttpContext.Current.Response.End();
                            }
                        }
                    }
                }
            }
        }
    }
}
