using System;
using NUnit.Framework;
using System.IO;
using OfficeOpenXml;

namespace XLSBuild.Test
{
	[TestFixture]
	public class TestXlsx
	{
		private const string TestExcelFile = "foo.xlsm";
		private const string TestVBAFile = "foo.bas";

		[SetUp]
		public void SetUp()
		{
			if (File.Exists (TestExcelFile)) {
				File.Delete (TestExcelFile);
			}
		}
		
		[Test]
		public void TestCreate(){
			var uut = new CreateXlsx (){Filename = TestExcelFile};
			Assert.That(uut.Execute (), Is.True);
			Assert.That (File.Exists (TestExcelFile), Is.True);
		}

        [Test]
        public void TestAddWorksheet()
        {
            CreateXls(TestExcelFile);
            var uut = new AddVbaModule()
            {
                Filename = TestExcelFile,
                Source = "../../foo.bas",
                Name = "foo",

            };

        }

        [Test]
        public void TestAddNamedMacro()
        {
            CreateXls(TestExcelFile);
            var uut = new AddVbaModule()
            {
                Filename = TestExcelFile,
                Source = "../../foo.bas",
                Name = "foo"
            };

            Assert.That(uut.Execute(), Is.True);
            Assert.That(File.Exists(TestExcelFile), Is.True);

            using (var package = new ExcelPackage(new FileInfo(TestExcelFile)))
            {
                Assert.That(package.Workbook.VbaProject, Is.Not.Null);
                Console.WriteLine(package.Workbook.VbaProject.Name);

                // assert that it contains module
                Assert.That(package.Workbook.VbaProject.Modules.Exists("foo"));
            }
        }

        [Test]
		public void TestAddMacro(){
			CreateXls (TestExcelFile);
			var uut = new AddVbaModule ()
            { 
				Filename = TestExcelFile,
				Source = "../../foo.bas",
			};

			Assert.That (uut.Execute (), Is.True);
			Assert.That (File.Exists (TestExcelFile), Is.True);

			using (var package = new ExcelPackage (new FileInfo (TestExcelFile))) 
            {
				Assert.That (package.Workbook.VbaProject, Is.Not.Null);
				Console.WriteLine (package.Workbook.VbaProject.Name);
                
                // assert that it contains module
                Assert.That(package.Workbook.VbaProject.Modules.Exists("Module 1"));
			}
		}

		private void CreateXls(string fileName){
			using (var package = new ExcelPackage (new System.IO.FileInfo (fileName))) 
            {
				package.Workbook.Worksheets.Add ("Sheet 1");
				package.Save ();
			}
		}
	}
}

