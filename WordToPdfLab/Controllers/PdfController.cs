using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using WordToPdfLab.Lib;

namespace WordToPdfLab.Controllers
{
	public class PdfController : Controller
	{
		// GET: Pdf
		public ActionResult Index()
		{
			return View();
		}

		public FileStreamResult Export()
		{
			byte[] result = null;
			Dictionary<string, string> fields = new Dictionary<string, string>();
			fields.Add("Name", "曾彥博");
			fields.Add("Job", "程式設計師");
			fields.Add("Year", "2018");
			fields.Add("Month", "11");
			fields.Add("Day", "26");

			//使用using確保Word資源被釋放
			using (var cvtr = new PdfConverter())
			{
				var buff = cvtr.GetPdf(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "templates\\Offer.docx"), fields);

				result = buff;
			}

			Response.AppendHeader("content-disposition", "inline; filename=Offer.pdf");
			return new FileStreamResult(new MemoryStream(result), "application/pdf");
		}
	}
}