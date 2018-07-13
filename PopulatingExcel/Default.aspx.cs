using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeOpenXml.Table;
using OfficeOpenXml;
using System.IO;

namespace PopulatingExcel
{
    public partial class Default : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void Populate_Click(object sender, EventArgs e)
        {
            string path = Server.MapPath(".") + "/Excel/template.xlsx";

            FileInfo newFile = new FileInfo(path.ToString());

            ExcelPackage pck = new ExcelPackage(newFile);
            var ws = pck.Workbook.Worksheets["Sheet1"];

            ws.Cells["O4"].Value = "A";
            ws.Cells["D12"].Value = "Some Name";
            ws.Cells["N12"].Value = "43534";
            ws.Cells["D14"].Value = "Last modified";
            ws.Cells["D16"].Style.Font.Bold = true;

            pck.Save();

            Response.ContentType = "application/excel";
            Response.AppendHeader("Content-Disposition", "attachment; filename=Sample.xlsx");
            Response.TransmitFile(Server.MapPath("/Excel/template.xlsx"));
            Response.End();
        }
    }
}