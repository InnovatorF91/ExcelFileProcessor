using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;

namespace ExcelFileProcessor
{
	class Program
	{
		static void Main(string[] args)
		{
			// Setting the console encoding to UTF-8
			Console.OutputEncoding = Encoding.UTF8;
			Console.InputEncoding = Encoding.UTF8;

			// Reading the configuration in App.config
			string folderPath = ConfigurationManager.AppSettings["FolderPath"];
			string keyword = ConfigurationManager.AppSettings["Keyword"];

			// Get all Excel files in the directory
			var excelFiles = Directory.EnumerateFiles(folderPath, "*.*", SearchOption.AllDirectories)
				.Where(s => s.EndsWith(".xlsx") || s.EndsWith(".csv") || s.EndsWith(".xlsm")).ToList();

			// Step 2: Read the files
			var textFolderPath = Path.Combine(folderPath, "TextFiles");

			if (!Directory.Exists(textFolderPath))
			{
				Directory.CreateDirectory(textFolderPath);

				foreach (var file in excelFiles)
				{
					var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(file);
					var textFilePath = Path.Combine(textFolderPath, fileNameWithoutExtension + ".txt");

					using (var textWriter = new StreamWriter(textFilePath))
					{
						if (file.EndsWith(".csv"))
						{
							foreach (var line in File.ReadLines(file))
							{
								textWriter.WriteLine(line);
							}
						}
						else
						{
							using (var stream = File.Open(file, FileMode.Open, FileAccess.Read))
							{
								using (var reader = ExcelReaderFactory.CreateReader(stream))
								{
									do
									{
										while (reader.Read())
										{
											for (int column = 0; column < reader.FieldCount; column++)
											{
												textWriter.Write(reader.GetValue(column)?.ToString() + "\t");
											}
											textWriter.WriteLine();
										}
									} while (reader.NextResult());
								}
							}
						}
					}
				}
			}

			var matchedFiles = new List<string>();

			var outputFile = "Output.txt";

			var outputFilePath = Path.Combine(textFolderPath, outputFile);

			foreach (var textFile in Directory.EnumerateFiles(textFolderPath, "*.txt"))
			{
				var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(textFile);
				var correspondingExcelFile = excelFiles.FirstOrDefault(f => Path.GetFileNameWithoutExtension(f) == fileNameWithoutExtension);

				if (File.ReadLines(textFile, Encoding.UTF8).Any(line => line.Contains(keyword)) && correspondingExcelFile != null)
				{
					matchedFiles.Add(correspondingExcelFile);
				}
			}

			// Output matched Excel files' names
			Console.WriteLine("以下は、キーワードを含むエクセルファイルの名前です：");
			foreach (var matchedFile in matchedFiles)
			{
				Console.WriteLine(matchedFile);
				File.AppendAllText(outputFilePath, matchedFile + Environment.NewLine);
			}

			Console.ReadKey();
		}
	}
}
