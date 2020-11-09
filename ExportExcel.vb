Imports System.Data.SqlClient
Imports System.IO
Imports System.Web.Configuration
Imports ClosedXML.Excel
Module ExportExcel
    Public Sub DoExport(FileName As String, SQL As String)


        Using con As New SqlConnection(ConnectionString)
            Using cmd As New SqlCommand(SQL)
                Using sda As New SqlDataAdapter()
                    cmd.Connection = con
                    sda.SelectCommand = cmd
                    Using dt As New DataTable()
                        sda.Fill(dt)
                        Using wb As New XLWorkbook()
        wb.Worksheets.Add(dt, "SheetName1")

                            HttpContext.Current.Response.Clear()
                            HttpContext.Current.Response.Buffer = True
                            HttpContext.Current.Response.Charset = ""
                            HttpContext.Current.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            HttpContext.Current.Response.AddHeader("content-disposition", "attachment;filename=" & FileName & ".xlsx")
                            Using MyMemoryStream As New MemoryStream()
                                wb.SaveAs(MyMemoryStream)
                                MyMemoryStream.WriteTo(HttpContext.Current.Response.OutputStream)
                                HttpContext.Current.Response.Flush()
                                HttpContext.Current.Response.End()
                            End Using
                        End Using
                    End Using
                End Using
            End Using
        End Using





    End Sub
End Module
