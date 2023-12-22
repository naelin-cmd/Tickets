using System;
using System.IO;
using Irony.Parsing;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using IronBarCode;
using Microsoft.VisualBasic;
using OfficeOpenXml.Drawing;
using System.Drawing;

namespace Register
{
    public class TicketV2
    {
        public string DmsOrderNumber { get; set; }
        public DateTime OrderDate { get; set; }
        public string DistributorName { get; set; }
        public int Isbn { get; set; }
        public string Title { get; set; }
        public string Quantity { get; set; }
        public string OrderDetails { get; set; }
        public string BookDetails { get; set; }

        static void Main(string[] args)
        {
            string dmsNumber="";
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            // Define the file paths for JSON input and Excel output
            string jsonFilePath = "Y:\\LSI\\DMS_ORDERS\\new_pod_orders_2023-12-20T09-53-34.json";
            string excelFilePath = "X:\\5. Users\\Naelin\\ticket\\ticketDispatch.xlsx"; // Modify the output file path as needed

            // Call the WriteTableToExcel method to save the data to an Excel file directly from JSON
            WriteTableToExcel(jsonFilePath, excelFilePath);
        }
     
        public static void WriteTableToExcel(string jsonFilePath, string excelFilePath)
        {
            using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                var worksheet = package.Workbook.Worksheets.Add("Ticket Data");

                // Set column headers
                worksheet.Cells[1, 1].Value = "DMS order number";
                worksheet.Cells[1, 2].Value = "Order date";
                worksheet.Cells[1, 3].Value = "Distributor name";
                worksheet.Cells[1, 5].Value = "Distributor Order Number";
                worksheet.Cells[1, 6].Value = "Barcode";
                //worksheet.Cells[1, 7].Value = "Despatch";
                worksheet.Cells[3, 1].Value = "ISBN";
                worksheet.Cells[3, 2].Value = "Title";
                worksheet.Cells[3, 3].Value = "Quantity";
                worksheet.Cells[1, 4].Value = "Address";
                worksheet.Cells[3, 5].Value = "Book Details";

                // Read and process JSON data directly
                using (var streamReader = new StreamReader(jsonFilePath))
                {
                    var json = streamReader.ReadToEnd();
                    var jsonObject = JObject.Parse(json);

                    // Make sure "orders" is a JArray
                    var ordersToken = jsonObject["orders"] as JArray;

                    if (ordersToken != null)
                    {
                        int row = 2; // Start writing data from the second row

                        foreach (var orderToken in ordersToken)
                        {
                            // Access "dms_order_number" within the current orderToken
                            var barcodeValue = orderToken["dms_order_number"]?.ToString() ?? "";
                            var myBarcode = BarcodeWriter.CreateBarcode(barcodeValue, BarcodeWriterEncoding.Code128);
                            myBarcode.ResizeTo(100, 25);
                            
                            var barcodeFilePath = $"Barcode_{row}.png"; // You can adjust the file format if needed
                            myBarcode.SaveAsPng(barcodeFilePath);

                            using (var excelImage = worksheet.Drawings.AddPicture($"Barcode_{row}", new FileInfo(barcodeFilePath)))
                            {
                                excelImage.SetPosition(row-1, 0, 5, 0); // Set the position of the image in the cell
                            }

                            // Delete the temporary barcode image file
                            

                            worksheet.Cells[row, 1].Value = orderToken["dms_order_number"]?.ToString() ?? "";
                            worksheet.Cells[row, 2].Value = orderToken["created_at"]?.ToString() ?? "";
                            worksheet.Cells[row, 3].Value = orderToken["distributor_name"]?.ToString() ?? "";
                            worksheet.Cells[row, 4].Value = orderToken["location"]?.ToString() ?? "";
                            worksheet.Cells[row, 5].Value = orderToken["distributor_order_number"]?.ToString() ?? "";
                           
                            // Access "order_details" within the current orderToken
                            var orderDetailsToken = orderToken["order_details"];
                            int orderDetailRow = row + 2;

                            int totalQuantity = 0; // Initialize total quantity for this order

                            foreach (var detailToken in orderDetailsToken)
                            {
                                int qty = Convert.ToInt32(detailToken["qty"]?.ToString() ?? "0");
                                totalQuantity += qty; // Add quantity to the total
                                worksheet.Cells[orderDetailRow, 3].Value = qty.ToString();
                                worksheet.Cells[orderDetailRow, 1].Value = detailToken["isbn"]?.ToString() ?? "";
                                worksheet.Cells[orderDetailRow, 2].Value = detailToken["title"]?.ToString() ?? "";
                                worksheet.Cells[orderDetailRow, 5].Value = $" Bind: {detailToken["binding"]}";









                                orderDetailRow++;
                            }




                            // Set the row variable to the next available row after processing order details
                            row = orderDetailRow;

                            // Add a row for the total quantity for this order
                            worksheet.Cells[row, 3].Value = totalQuantity;

                            // Move to the next row for the next order
                            row++;
                        }
                    }



                    var usedRange = worksheet.Cells[worksheet.Dimension.Address];
                    usedRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    usedRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    usedRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    usedRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                    worksheet.HeaderFooter.OddFooter.CenteredText = "&P of &N";

                    package.Save();
                }
            }
        }
    }
}
