using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.ComponentModel;
using System.Drawing;
using System.Web;

namespace WebApplication1
{
    /// <summary>
    /// HTTP Handler for generating Excel template for hotel bookings
    /// </summary>
    public class DownloadTemplate : IHttpHandler
    {
        public void ProcessRequest(HttpContext context)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            try
            {
                using (ExcelPackage package = new ExcelPackage())
                {
                    // Create worksheet
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Hotel Bookings");

                    // Set column headers
                    string[] headers = new string[]
                    {
                        "Guest Name",
                        "Guest Email",
                        "Guest Phone",
                        "Hotel Code",
                        "Room Type",
                        "Check-In Date",
                        "Check-Out Date",
                        "Number of Guests",
                        "Special Requests"
                    };

                    // Style headers
                    for (int i = 0; i < headers.Length; i++)
                    {
                        var cell = worksheet.Cells[1, i + 1];
                        cell.Value = headers[i];
                        cell.Style.Font.Bold = true;
                        cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        cell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189));
                        cell.Style.Font.Color.SetColor(Color.White);
                        cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        cell.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    }

                    // Add sample data rows
                    AddSampleRow(worksheet, 2, "John Doe", "john.doe@email.com", "+1-555-0123",
                        "HTL001", "Deluxe", "2025-12-15", "2025-12-18", "2", "Late check-in preferred");

                    AddSampleRow(worksheet, 3, "Jane Smith", "jane.smith@email.com", "+1-555-0456",
                        "HTL002", "Suite", "2025-12-20", "2025-12-25", "4", "Ocean view room");

                    AddSampleRow(worksheet, 4, "Bob Johnson", "bob.j@email.com", "+1-555-0789",
                        "HTL003", "Standard", "2026-01-05", "2026-01-10", "2", "");

                    // Add instructions sheet
                    ExcelWorksheet instructionsSheet = package.Workbook.Worksheets.Add("Instructions");
                    AddInstructions(instructionsSheet);

                    // Add hotel codes reference sheet
                    ExcelWorksheet hotelCodesSheet = package.Workbook.Worksheets.Add("Hotel Codes");
                    AddHotelCodesReference(hotelCodesSheet);

                    // Add room types reference sheet
                    ExcelWorksheet roomTypesSheet = package.Workbook.Worksheets.Add("Room Types");
                    AddRoomTypesReference(roomTypesSheet);

                    // Auto-fit columns
                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                    instructionsSheet.Cells[instructionsSheet.Dimension.Address].AutoFitColumns();
                    hotelCodesSheet.Cells[hotelCodesSheet.Dimension.Address].AutoFitColumns();
                    roomTypesSheet.Cells[roomTypesSheet.Dimension.Address].AutoFitColumns();

                    // Freeze header row
                    worksheet.View.FreezePanes(2, 1);

                    // Set response headers
                    context.Response.Clear();
                    context.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    context.Response.AddHeader("Content-Disposition",
                        "attachment; filename=Hotel_Booking_Template_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx");

                    // Write to response
                    package.SaveAs(context.Response.OutputStream);
                    context.Response.End();
                }
            }
            catch (Exception ex)
            {
                context.Response.ContentType = "text/plain";
                context.Response.Write("Error generating template: " + ex.Message);
            }
        }

        private void AddSampleRow(ExcelWorksheet worksheet, int row, string guestName,
            string email, string phone, string hotelCode, string roomType,
            string checkIn, string checkOut, string guests, string requests)
        {
            worksheet.Cells[row, 1].Value = guestName;
            worksheet.Cells[row, 2].Value = email;
            worksheet.Cells[row, 3].Value = phone;
            worksheet.Cells[row, 4].Value = hotelCode;
            worksheet.Cells[row, 5].Value = roomType;
            worksheet.Cells[row, 6].Value = checkIn;
            worksheet.Cells[row, 7].Value = checkOut;
            worksheet.Cells[row, 8].Value = guests;
            worksheet.Cells[row, 9].Value = requests;

            // Apply light gray background to sample data
            for (int col = 1; col <= 9; col++)
            {
                var cell = worksheet.Cells[row, col];
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(242, 242, 242));
                cell.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            }
        }

        private void AddInstructions(ExcelWorksheet worksheet)
        {
            int row = 1;

            // Title
            worksheet.Cells[row, 1].Value = "HOTEL BOOKING BULK UPLOAD - INSTRUCTIONS";
            worksheet.Cells[row, 1].Style.Font.Bold = true;
            worksheet.Cells[row, 1].Style.Font.Size = 16;
            worksheet.Cells[row, 1].Style.Font.Color.SetColor(Color.FromArgb(79, 129, 189));
            row += 2;

            // Instructions
            AddInstructionLine(worksheet, ref row, "1. GENERAL GUIDELINES:");
            AddInstructionLine(worksheet, ref row, "   • Fill in all required fields (marked with * in column headers)");
            AddInstructionLine(worksheet, ref row, "   • Delete the sample data rows before uploading");
            AddInstructionLine(worksheet, ref row, "   • Do not modify the column headers");
            AddInstructionLine(worksheet, ref row, "   • Maximum 1000 records per upload");
            row++;

            AddInstructionLine(worksheet, ref row, "2. FIELD SPECIFICATIONS:");
            row++;

            AddInstructionLine(worksheet, ref row, "   Guest Name*:");
            AddInstructionLine(worksheet, ref row, "   • Full name of the guest (2-200 characters)");
            row++;

            AddInstructionLine(worksheet, ref row, "   Guest Email*:");
            AddInstructionLine(worksheet, ref row, "   • Valid email address format (e.g., user@domain.com)");
            row++;

            AddInstructionLine(worksheet, ref row, "   Guest Phone*:");
            AddInstructionLine(worksheet, ref row, "   • Contact number with country code (e.g., +1-555-1234)");
            row++;

            AddInstructionLine(worksheet, ref row, "   Hotel Code*:");
            AddInstructionLine(worksheet, ref row, "   • Valid hotel code from 'Hotel Codes' sheet");
            AddInstructionLine(worksheet, ref row, "   • Examples: HTL001, HTL002, HTL003");
            row++;

            AddInstructionLine(worksheet, ref row, "   Room Type*:");
            AddInstructionLine(worksheet, ref row, "   • Valid room type from 'Room Types' sheet");
            AddInstructionLine(worksheet, ref row, "   • Examples: Standard, Deluxe, Suite, Family Room, Executive");
            row++;

            AddInstructionLine(worksheet, ref row, "   Check-In Date*:");
            AddInstructionLine(worksheet, ref row, "   • Format: YYYY-MM-DD or MM/DD/YYYY");
            AddInstructionLine(worksheet, ref row, "   • Must be today or future date");
            row++;

            AddInstructionLine(worksheet, ref row, "   Check-Out Date*:");
            AddInstructionLine(worksheet, ref row, "   • Format: YYYY-MM-DD or MM/DD/YYYY");
            AddInstructionLine(worksheet, ref row, "   • Must be after Check-In Date");
            row++;

            AddInstructionLine(worksheet, ref row, "   Number of Guests*:");
            AddInstructionLine(worksheet, ref row, "   • Numeric value (1-10)");
            row++;

            AddInstructionLine(worksheet, ref row, "   Special Requests:");
            AddInstructionLine(worksheet, ref row, "   • Optional field for additional requirements");
            AddInstructionLine(worksheet, ref row, "   • Maximum 1000 characters");
            row += 2;

            AddInstructionLine(worksheet, ref row, "3. VALIDATION RULES:");
            AddInstructionLine(worksheet, ref row, "   • All dates must be valid and logical");
            AddInstructionLine(worksheet, ref row, "   • Email addresses must be unique per booking");
            AddInstructionLine(worksheet, ref row, "   • Hotel codes and room types must exist in the system");
            AddInstructionLine(worksheet, ref row, "   • Check-out date must be after check-in date");
            row += 2;

            AddInstructionLine(worksheet, ref row, "4. SUPPORT:");
            AddInstructionLine(worksheet, ref row, "   • For assistance, contact: booking-support@hotelchain.com");
            AddInstructionLine(worksheet, ref row, "   • Phone: +1-800-HOTEL-HELP");

            worksheet.Column(1).Width = 100;
        }

        private void AddInstructionLine(ExcelWorksheet worksheet, ref int row, string text)
        {
            worksheet.Cells[row, 1].Value = text;
            worksheet.Cells[row, 1].Style.WrapText = true;
            row++;
        }

        private void AddHotelCodesReference(ExcelWorksheet worksheet)
        {
            // Headers
            worksheet.Cells[1, 1].Value = "Hotel Code";
            worksheet.Cells[1, 2].Value = "Hotel Name";
            worksheet.Cells[1, 3].Value = "Location";
            worksheet.Cells[1, 4].Value = "City";
            worksheet.Cells[1, 5].Value = "Country";

            for (int i = 1; i <= 5; i++)
            {
                worksheet.Cells[1, i].Style.Font.Bold = true;
                worksheet.Cells[1, i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[1, i].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189));
                worksheet.Cells[1, i].Style.Font.Color.SetColor(Color.White);
            }

            // Sample hotel data
            AddHotelCode(worksheet, 2, "HTL001", "Grand Plaza Hotel", "Marina Bay", "Singapore", "Singapore");
            AddHotelCode(worksheet, 3, "HTL002", "Sunset Resort & Spa", "Bali Beach", "Denpasar", "Indonesia");
            AddHotelCode(worksheet, 4, "HTL003", "Mountain View Inn", "Shimla Hills", "Shimla", "India");
            AddHotelCode(worksheet, 5, "HTL004", "City Central Hotel", "Downtown", "Bangkok", "Thailand");
            AddHotelCode(worksheet, 6, "HTL005", "Riverside Retreat", "Riverside District", "Mandalay", "Myanmar");

            worksheet.View.FreezePanes(2, 1);
        }

        private void AddHotelCode(ExcelWorksheet worksheet, int row, string code,
            string name, string location, string city, string country)
        {
            worksheet.Cells[row, 1].Value = code;
            worksheet.Cells[row, 2].Value = name;
            worksheet.Cells[row, 3].Value = location;
            worksheet.Cells[row, 4].Value = city;
            worksheet.Cells[row, 5].Value = country;
        }

        private void AddRoomTypesReference(ExcelWorksheet worksheet)
        {
            // Headers
            worksheet.Cells[1, 1].Value = "Room Type";
            worksheet.Cells[1, 2].Value = "Description";
            worksheet.Cells[1, 3].Value = "Max Occupancy";

            for (int i = 1; i <= 3; i++)
            {
                worksheet.Cells[1, i].Style.Font.Bold = true;
                worksheet.Cells[1, i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[1, i].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189));
                worksheet.Cells[1, i].Style.Font.Color.SetColor(Color.White);
            }

            // Room types
            AddRoomType(worksheet, 2, "Standard", "Comfortable room with basic amenities", "2");
            AddRoomType(worksheet, 3, "Deluxe", "Spacious room with premium amenities", "2");
            AddRoomType(worksheet, 4, "Suite", "Luxury suite with separate living area", "4");
            AddRoomType(worksheet, 5, "Family Room", "Large room suitable for families", "4");
            AddRoomType(worksheet, 6, "Executive", "Business-class room with workspace", "2");

            worksheet.View.FreezePanes(2, 1);
        }

        private void AddRoomType(ExcelWorksheet worksheet, int row, string type, string desc, string maxOcc)
        {
            worksheet.Cells[row, 1].Value = type;
            worksheet.Cells[row, 2].Value = desc;
            worksheet.Cells[row, 3].Value = maxOcc;
        }

        public bool IsReusable
        {
            get { return false; }
        }
    }
}