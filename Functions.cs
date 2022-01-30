using System;
using System.IO;
using ExcelDna.Integration;
using Microsoft.Data.Sqlite;

namespace UsingMicrosoftDataSqlite
{
    public static class MyFunctions
    {
        static SqliteConnection _connection;
        static SqliteCommand _productNameCommand;

        private static void EnsureConnection()
        {
            if (_connection == null)
            {
                var xllDirectory = Path.GetDirectoryName(ExcelDnaUtil.XllPath);
                var dbPath = Path.Combine(xllDirectory, @"Data\Northwind.db");
                _connection = new SqliteConnection($"Data Source={dbPath}");
                _connection.Open();

                _productNameCommand = new SqliteCommand("SELECT ProductName FROM Products WHERE ProductID = @ProductID", _connection);
                _productNameCommand.Parameters.Add("@ProductID", SqliteType.Integer);
            }
        }

        public static object ProductName(int productID)
        {
            try
            {
                EnsureConnection();
                _productNameCommand.Parameters["@ProductID"].Value = productID;
                return _productNameCommand.ExecuteScalar();
            }
            catch (Exception ex)
            {
                return ex.ToString();
            }
        }

    }

}
