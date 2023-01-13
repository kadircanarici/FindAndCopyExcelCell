using System;
using System.Data;
using OfficeOpenXml.FormulaParsing.Excel;
using ExcelDataReader;
using ExcelDataWriter;

using ClosedXML.Excel;
using System.IO;
using System.Text;

class Program
{
    static void Main(string[] args)
    {
        Console.WriteLine("Lütfen ilk excel dosyasının yolunu giriniz (Sheet1):");
        string file1Path = Console.ReadLine();
        // Kullanicidan ikinci excel dosyasinin yolunu al
        Console.WriteLine("Lütfen ikinci excel dosyasının yolunu giriniz (Sheet2):");
        string file2Path = Console.ReadLine();
       

        // Excel dosyalarını okuma
        DataSet data1 = ReadExcelFile(file1Path);
        DataSet data2 = ReadExcelFile(file2Path);

        // İlk excel dosyasının ilk kolonu
        DataTable data1Col1 = data1.Tables[0];

        // İkinci excel dosyasının ilk kolonu
        DataTable data2Col1 = data2.Tables[0];

        // İlk excel dosyasının ilk kolonunda her satır için
        foreach (DataRow row1 in data1Col1.Rows)
        {
            string val1 = row1[0].ToString();
            Console.WriteLine(val1);

            // İkinci excel dosyasındaki ilk kolonunda satırları arama
            foreach (DataRow row2 in data2Col1.Rows)
            {
                string val2 = row2[0].ToString();
                Console.WriteLine(val2);

                // Eşleşme bulursa
                if (val1 == val2)
                {
                    Console.WriteLine(val1+ " " + val2);
                    // İlk excel dosyasının ikinci kolonundaki değeri 
                    // ikinci excel dosyasında eşleştiği satırın ikinci kolonuna yaz
                    row2[1] = row1[1];

                }
            }
        }

        // İkinci excel dosyasını kaydet
        SaveExcelFile(file2Path, data2);
        Console.ReadKey();
    }

    static DataSet ReadExcelFile(string filePath)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
        {
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration()
                    {
                        UseHeaderRow = true
                    }
                });
                // code to process the data in the DataSet
                //IExcelDataReader readera = ExcelReaderFactory.CreateOpenXmlReader(File.OpenRead(filePath));
                DataSet data = reader.AsDataSet();
                reader.Close();
                return data;
            }
        }
        // Excel dosyasını okuma
       
    }



static void SaveExcelFile(string filePath, DataSet data)
{
    using (var workbook = new XLWorkbook())
    {
        foreach (DataTable table in data.Tables)
        {
            var worksheet = workbook.Worksheets.Add(table.TableName);
            worksheet.Cell(1, 1).InsertTable(table);
        }
        workbook.SaveAs(filePath);
    }
}

}
