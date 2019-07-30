using Syncfusion.XlsIO;
using System.Collections.Generic;

namespace TemplateMarkers
{
    class Program
    {
        static void Main(string[] args)
        {
            // Init rows collection.
            var records = new List<ReportRow>()
            {
                new ReportRow() {Name = "Product 1", Quantity = 2, Price = 100},
                new ReportRow() {Name = "Product 2", Quantity = 1, Price = 200},
                new ReportRow() {Name = "Product 3", Quantity = 5, Price = 300},
                new ReportRow() {Name = "Product 4", Quantity = 10, Price = 150},
                new ReportRow() {Name = "Product 5", Quantity = 7, Price = 100}
            };

            //Creates a new instance for ExcelEngine
            ExcelEngine excelEngine = new ExcelEngine();

            //Loads or open an existing workbook through Open method of IWorkbooks
            var inputFileName = @"Template.xlsx";
            var resultFileName = @"Report.xlsx";

            IWorkbook workbook = excelEngine.Excel.Workbooks.Open(inputFileName);

            //Create Template Marker Processor
            ITemplateMarkersProcessor marker = workbook.CreateTemplateMarkersProcessor();

            //Add collections to the marker variables where the name should match with input template
            marker.AddVariable("records", records);

            //Process the markers in the template
            marker.ApplyMarkers();

            workbook.SaveAs(resultFileName);


        }
    }
}
