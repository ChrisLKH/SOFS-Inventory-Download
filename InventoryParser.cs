using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;

namespace SOFSInventoryDownloader
{
    public static class InventoryParser
    {
        public static DataTable Parse(string excelFilePath)
        {
            var sourceTable = ExcelHelper.GetDataTableFromExcel(excelFilePath);
            var inventoryTable = new DataTable();

            inventoryTable.Columns.AddRange(new[]
            {
                new DataColumn("SupplierSKU"),
                new DataColumn("Sellable", typeof(int)),
                new DataColumn("KSQty", typeof(int)),
                new DataColumn("SLCQty", typeof(int)),
                new DataColumn("CPAQty", typeof(int)),
                new DataColumn("FCAQty", typeof(int)),
                new DataColumn("UPC")
            });

            foreach (DataRow row in sourceTable.Rows)
            {
                int KSQty = 0;
                int SLCQty = GetNetQty(row, "Salt Lake City, UT");
                int CPAQty = GetNetQty(row, "Carlisle, PA");
                int FCAQty = GetNetQty(row, "Fontana, CA");

                int sellable = KSQty + SLCQty + CPAQty + FCAQty;

                var newRow = inventoryTable.NewRow();
                newRow["SupplierSKU"] = row["Supplier SKU"];
                newRow["Sellable"] = sellable;
                newRow["KSQty"] = KSQty;
                newRow["SLCQty"] = SLCQty;
                newRow["CPAQty"] = CPAQty;
                newRow["FCAQty"] = FCAQty;
                newRow["UPC"] = row["UPC"];
                inventoryTable.Rows.Add(newRow);
            }

            return inventoryTable;
        }

        public static void SaveToDatabase(DataTable inventory)
        {
            string connString = "Server=localhost;Database=master;User Id=user;Password=password;";
            using var conn = new SqlConnection(connString);
            conn.Open();

            using (var deleteCmd = new SqlCommand("DELETE FROM [SofsInventory]", conn))
            {
                deleteCmd.ExecuteNonQuery();
            }

            using var bulkCopy = new SqlBulkCopy(conn)
            {
                DestinationTableName = "[SofsInventory]"
            };
            bulkCopy.WriteToServer(inventory);
        }

        public static void ArchiveFile(string path)
        {
            string timestampedPath = $@"C:\\SOFS_Inventory_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
            string networkPath = @"\\server\SOFS Inventory\Inventory_Summary.xlsx";

            File.Copy(path, timestampedPath);
            File.Copy(path, networkPath, true);
            File.Delete(path);
        }

        private static int GetNetQty(DataRow row, string warehouse)
        {
            int onHand = int.Parse(row[$"{warehouse} Inventory on hand"].ToString());
            int reserved = int.Parse(row[$"{warehouse} Reserved Inventory"].ToString());
            int onHold = int.Parse(row[$"{warehouse} Inventory on hold"].ToString());
            return onHand - reserved - onHold;
        }
    }
}
