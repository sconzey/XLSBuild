using System;
using Microsoft.Build.Framework;
using Microsoft.Build.Utilities;
using OfficeOpenXml;
using OfficeOpenXml.VBA;
using System.IO;

namespace XLSBuild
{
	public class AddVbaModule : Task
	{
        private enum VbaModuleType
        {
            Module = 0,
            Worksheet = 1,
            ThisWorksheet = 2
        }

        public AddVbaModule()
        {
            Type = VbaModuleType.Module.ToString();
        }

        private VbaModuleType GetModuleType()
        {
            VbaModuleType type;
            if (!Enum.TryParse<VbaModuleType>(Type, out type))
                throw new Exception(string.Format("Module type not supported: {0}. It must be one of: {1}", 
                    string.Join(", ", Enum.GetNames(typeof(VbaModuleType))), Type));
            return type;
        }

		public override bool Execute()
		{
			using(var package = new ExcelPackage(new FileInfo(Filename)))
            {
				if (package.Workbook.VbaProject == null) {
					package.Workbook.CreateVBAProject ();
				}
				var proj = package.Workbook.VbaProject;
                
                var name = Name ?? FindDefaultName(proj.Modules);
				var mod = proj.Modules.AddModule(name);
				mod.Code = File.ReadAllText (Source);

				package.Save ();
			}
			return true;
		}

        private string FindDefaultName(ExcelVbaModuleCollection modules)
        {
            int i = 1;
            string moduleName;
            while(modules.Exists(moduleName = string.Format("Module {0}", i))) 
                i++;
            return moduleName;
        }

        [Required]
		public string Filename { get; set; }

		[Required]
        public string Source { get; set; }
		
		public string Name { get; set; }

        public string Type { get; set; }
	}
}
