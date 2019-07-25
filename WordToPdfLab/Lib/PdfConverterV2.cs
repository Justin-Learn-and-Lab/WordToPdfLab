using Spire.Doc;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Web;

namespace WordToPdfLab.Lib
{
	public class PdfConverterV2
	{
		public PdfConverterV2()
		{
			
		}

		public void GetPdf(string templateFile, Dictionary<string, string> fields)
		{
			object filePath = templateFile;

			Stream stream = null;

			//檔案先寫入系統暫存目錄
			object outFile = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pdf");

			try
			{
				using (Document document = new Document())
				{
					document.LoadFromFile(templateFile);

					foreach (var item in fields)
					{
						document.Replace("[$$" + item.Key + "$$]", item.Value, false, false);
					}

					document.SaveToFile(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "templates\\temp.pdf"), FileFormat.PDF);

					document.SaveToFile(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "templates\\temp.doc"), FileFormat.Doc);
				}
			}
			catch (Exception ex)
			{
				Debug.WriteLine(ex.ToString());
			}
		}
	}
}