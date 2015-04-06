using System;
using Microsoft.Build.Framework;
using Microsoft.Build.Utilities;
using OfficeOpenXml;
using System.IO;

namespace XLSBuild
{

	public class AddVbaModule : Task
	{
		public override bool Execute()
		{
			using(var package = new ExcelPackage(new FileInfo(Filename))){
				if (package.Workbook.VbaProject == null) {
					package.Workbook.CreateVBAProject ();
				}
				var proj = package.Workbook.VbaProject;
				var mod = proj.Modules.AddModule(Name);
				mod.Code = File.ReadAllText (Source);
				package.Save ();
			}
			return true;
		}

		public string Filename { get; set; }
		public string Source { get; set; }
		// TODO: default to sensible name
		public string Name { get; set; }
		// TODO: handle THisWorksheet code
	}

}
