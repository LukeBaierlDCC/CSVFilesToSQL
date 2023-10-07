using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Configuration;

namespace CSVFile
{
    class Program
    {
        static void Main(string[] args)
        {
            string connectionString = "Server=.;Database=myDataBase;Integrated Security=True;";
            string filePath = @"C:\Users\lukee\Documents\WASD\WASD FM world\DreamWASDFMReferences.ods";
            string tableName = "myTable";

            // Create table if it doesn't exist
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                try
                {
                    con.Open();
                    using (SqlCommand command = new SqlCommand($"IF OBJECT_ID('{tableName}', 'U') IS NULL CREATE TABLE {tableName} (ID int, VideoGame varchar(55), DebutYear int, Developer varchar(100), Publisher varchar(100), Genre varchar(55), Systems varchar(255), NotableCharacters varchar(255))", con))
                    {
                        command.ExecuteNonQuery();
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    Console.ReadLine();
                }
            }

            // Read CSV file and insert into SQL Server table
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                try
                {
                    con.Open();
                    using (SqlCommand command = new SqlCommand($"BULK INSERT {tableName} FROM '{filePath}' WITH (FORMAT='CSV', FIRSTROW=2)", con))
                    {
                        command.ExecuteNonQuery();
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    Console.ReadLine();
                }
            }

            Console.WriteLine("Data has been successfully imported from CSV to SQL Server.");
            Console.ReadLine();
        }
    }
}
