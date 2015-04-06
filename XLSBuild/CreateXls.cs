using System;
using Microsoft.Build.Framework;
using Microsoft.Build.Utilities;
using OfficeOpenXml;
using System.IO;

namespace XLSBuild
{
	public class CreateXlsx : Task
	{
		public override bool Execute()
		{
			using(var package = new ExcelPackage(new FileInfo(Filename))){
				package.Workbook.Worksheets.Add ("Sheet 1");
				package.Save ();
			}
			return true;
		}

        [Required]
		public string Filename { get; set; }
	}
}

