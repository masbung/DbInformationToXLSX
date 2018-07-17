using DbToXLSX.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace DbToXLSX.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index() {
            return View();
        }

        public ActionResult About() {            
            return View();
        }

        public ActionResult Contact() {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        public FileResult DownloadCSV() {

            string connectionString = ConfigurationManager.ConnectionStrings["DefaultDatabase"].ConnectionString;
            string dbName = ConfigurationManager.AppSettings["DatabaseName"];

            var model = new HomeViewModel();

            var selectTables = String.Format("SELECT TABLE_SCHEMA, TABLE_NAME"
                + " FROM INFORMATION_SCHEMA.TABLES"
                + " WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_CATALOG='{0}'"                
                + " ORDER BY TABLE_SCHEMA", dbName);

            using (SqlConnection connection = new SqlConnection(connectionString)) {
                connection.Open();

                using (SqlCommand command = new SqlCommand(selectTables, connection)) {
                    using (SqlDataReader reader = command.ExecuteReader()) {
                        while (reader.Read()) {
                            model.Tables.Add(new Table {
                                SchemaName = reader.GetString(reader.GetOrdinal("TABLE_SCHEMA")),
                                TableName = reader.GetString(reader.GetOrdinal("TABLE_NAME")),
                            });
                        }
                    }
                }
            }

            var schemaVM = new SchemaViewModel();

            Schema schema = null;

            for (int i = 0; i < model.Tables.Count; i++) {
                var item = model.Tables[i];

                if (i == 0) {
                    schema = new Schema {
                        Name = item.SchemaName
                    };
                    schemaVM.Schemas.Add(schema);
                } else {
                    if (!schema.Name.Equals(item.SchemaName)) {
                        schema = new Schema {
                            Name = item.SchemaName
                        };
                        schemaVM.Schemas.Add(schema);
                    }
                }

                schema.Tables.Add(new Table {
                    Print = item.Print,
                    TableName = item.TableName,
                    SchemaName = item.SchemaName
                });
            }

            var excelPackage = this.GenerateExcel(schemaVM);
            Byte[] bin = excelPackage.GetAsByteArray();
            string file = "Tables.xlsx";

            return File(bin, System.Net.Mime.MediaTypeNames.Application.Octet, file);
        }

        private ExcelPackage GenerateExcel(SchemaViewModel schemaVM) {
            var ep = new ExcelPackage();

            foreach (var item in schemaVM.Schemas) {
                this.CreateSheet(ep, item);
            }

            return ep;
        }

        private string TableInfo(string dbName, string schemaName, string tableName) {
            return String.Format(
                " SELECT COLUMN_NAME, DATA_TYPE, IS_NULLABLE"
                + " FROM [{0}].INFORMATION_SCHEMA.COLUMNS"
                + " WHERE TABLE_SCHEMA = '{1}' AND TABLE_NAME = N'{2}'", dbName, schemaName, tableName);
        }

        private void CreateSheet(ExcelPackage ep, Schema schema) {
            string connectionString = ConfigurationManager.ConnectionStrings["DefaultDatabase"].ConnectionString;
            string dbName = ConfigurationManager.AppSettings["DatabaseName"];

            ep.Workbook.Worksheets.Add(schema.Name);
            var ws = ep.Workbook.Worksheets.Where(s => s.Name.Equals(schema.Name)).FirstOrDefault();

            var horizontalAligment = ExcelHorizontalAlignment.Center;
            var verticalAligment = ExcelVerticalAlignment.Center;

            var rowSchemaAndTableOne = 1;
            var rowInformationOne = 2;

            if (ws != null) {
                ws.Cells.Style.Font.Size = 11; //Default font size for whole sheet
                ws.Cells.Style.Font.Name = "Calibri"; //Default Font name for whole sheet

                //ExcelRange schemaCell = null;                
                ExcelRange tableCell = null;

                ExcelRange columnValueCellOne = null;
                ExcelRange dataTypeValueCellOne = null;
                ExcelRange nullableValueCellOne = null;

                using (SqlConnection connection = new SqlConnection(connectionString)) {
                    connection.Open();

                    foreach (var table in schema.Tables) {                        
                        tableCell = ws.Cells[rowSchemaAndTableOne, 1, rowSchemaAndTableOne, 2];
                        var tableCellBorder = tableCell.Style.Border;
                        tableCell.Merge = true;
                        tableCell.Style.Font.Bold = true;
                        tableCell.Value = String.Format("[{0}].[{1}]", schema.Name, table.TableName);
                        tableCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        tableCell.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                        tableCellBorder.Bottom.Style = tableCellBorder.Top.Style = tableCellBorder.Left.Style = tableCellBorder.Right.Style = ExcelBorderStyle.Thin;
                        tableCell.Style.HorizontalAlignment = horizontalAligment;
                        tableCell.Style.VerticalAlignment = verticalAligment;

                        using (SqlCommand command = new SqlCommand(this.TableInfo(dbName, schema.Name, table.TableName), connection)) {
                            using (SqlDataReader reader = command.ExecuteReader()) {
                                while (reader.Read()) {
                                    columnValueCellOne = ws.Cells[rowInformationOne, 1];
                                    columnValueCellOne.Value = reader.GetString(reader.GetOrdinal("COLUMN_NAME"));
                                    dataTypeValueCellOne = ws.Cells[rowInformationOne, 2];
                                    dataTypeValueCellOne.Value = reader.GetString(reader.GetOrdinal("DATA_TYPE"));
                                    nullableValueCellOne = ws.Cells[rowInformationOne, 3];
                                    nullableValueCellOne.Value = reader.GetString(reader.GetOrdinal("IS_NULLABLE"));
                                    rowInformationOne++;
                                }
                            }
                        }

                        rowInformationOne++;
                        rowSchemaAndTableOne = rowInformationOne;
                        rowInformationOne++;
                    }
                }
            }
        }
    }
}