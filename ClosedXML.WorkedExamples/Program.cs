using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.IO;
using System.IO.Packaging;



namespace ClosedXML.WorkedExamples
{
    public class Program
    {

        public static string BaseCreatedDirectory
        {
            get
            {
                //var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Created");
                var path = Path.Combine(Directory.GetCurrentDirectory(), "output/Created");
                if (!Directory.Exists(path)) Directory.CreateDirectory(path);
                return path;
            }
        }

        public static string BaseModifiedDirectory
        {
            get
            {
                //var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Modified");
                var path = Path.Combine(Directory.GetCurrentDirectory(), "output/Modified");
                if (!Directory.Exists(path)) Directory.CreateDirectory(path);
                return path;
            }
        }

        static void Main(string[] args)
        {
            var path = Program.BaseCreatedDirectory;

            var filePath1 = Path.Combine(path, "Formulae.xlsx");
            CreateTestDocument(filePath1, false);
            UnpackPackage(filePath1);

            var filePath2 = Path.Combine(path, "FormulaeWithIteration.xlsx");
            CreateTestDocument(filePath2, true);
            UnpackPackage(filePath2);

            var cwd = Directory.GetCurrentDirectory();
            var filePath3 = PathCombine(cwd, "/ClosedXML.WorkedExamples/data/test-with-iteration.xlsx");
            UnpackPackage(filePath3, Path.Combine(path, "test-with-iteration.xlsx-unpacked"));

        } // static void Main(string[] args)


        /// <summary>
        /// make sure the returned string has only directory separators valid on the current platform
        /// </summary>
        /// <remarks>
        /// <para>As a string representing a path segment</para>
        /// </remarks>
        public static string PathAdjustSeparators(string path)
        {
            string newPath =  path.Replace('/', Path.DirectorySeparatorChar);
            newPath = newPath.Replace('\\', Path.DirectorySeparatorChar);
            return newPath;
        }

        /// <summary>
        /// replacement for Path.Combine that should work with mixed path separators
        /// </summary>
        /// <remarks>
        /// <para>As a string representin a path segment</para>
        /// <para>As a string representin a path segment</para>
        /// </remarks>
        public static string PathCombine(string first, string second)
        {
            string path1 = PathAdjustSeparators(first);
            string path2 = PathAdjustSeparators(second);
            // For Path.Cobine to work, the second path segment MUST NOT start with a DirectorySeparatorChar!
            if (path2.StartsWith(Path.DirectorySeparatorChar))
            {
                path2 = path2.Substring(1);
            }
            return Path.Combine(path1, path2);
        }

        public static void CreateTestDocument(string fileName, bool withIteration = false)
        {
            if (File.Exists(fileName))
            {
                File.Delete(fileName);
            }

            string sheetName = withIteration ? "Formulae With Iteration" : "Formulae";

            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add(sheetName);
            ws.Cell("A1").Value = "Value 1";
            ws.Cell("B1").Value = "Value 2";
            ws.Cell("C1").Value = "Product";
            ws.Cell("E1").Value = "Iteration";

            ws.Cell("A3").Value = 1.2345;
            ws.Cell("B3").Value = 2.3456;
            ws.Cell("C3").FormulaA1 = "=A3*B3";
            ws.Cell("A4").FormulaA1 = "=IF(A2>1.1, 1.1, 2.2*B3)";

            ws.Cell("D3").FormulaA1 = "=IF(E3=0, 0.001, E3)";
            ws.Cell("D5").FormulaA1 = "=IF(E5=0, 0.1, E5)";

            if (withIteration)
            {
                // iteration with cells D3 and E3:   
                ws.Cell("E3").FormulaA1 = "=(0.8+D3)*C3";
                // iteration with cells D5 and E5:
                ws.Cell("E5").FormulaA1 = "=2.3-D5";
                // enable and configure iteration:
                wb.Iterate = true; // Default is flase (isn't it?)
                wb.IterateCount = 150; // Default is 100
                wb.IterateDelta = 0.01; // Default is 0.001
                // And the workbook calculation mode:
                wb.CalculateMode = XLCalculateMode.Auto;
                wb.FullCalculationOnLoad = true;
            }
            Console.WriteLine("saving XLWorkbook as {0}", fileName);
            wb.SaveAs(fileName);
        } // public static void CreateTestDocument(string fileName, bool withIteration = false)



        public static void UnpackPackage(
            string filePath,
            string targetDirectory = ""
        )
        {
            // open the package for reading
            using (
                Package package =
                    Package.Open(filePath, FileMode.Open, FileAccess.Read)
            )
            {
                if (targetDirectory.Length < 1)
                {
                    targetDirectory = filePath + "-unpacked";
                }

                // unpack the package to target directory
                UnpackPackage(package, targetDirectory);

                // close the package
                package.Close();
            } // using ...
        } // static public void UnpackPackage(string filePath, string targetDirectory = "")

        // unpack the given package to the filesystem in the given directory and format the XML parts nicely
        // the directory and the required subdirectories will be created
        // the given packagg is not modified (you may pass a read-only file)
        public static void UnpackPackage(Package package, string targetDirectory)
        {
            // create the target directory
            CreateDirectory(targetDirectory);

            // get all package parts contained in the package
            PackagePartCollection packageParts = package.GetParts();

            // loop over the package's parts and process each part
            foreach (PackagePart packagePart in packageParts)
            {
                Uri uri = packagePart.Uri;
                Console.WriteLine("Package part: {0}", uri);

                // construct a file name:
                string fileName = targetDirectory + uri;
                string dirName = Path.GetDirectoryName(fileName);
                CreateDirectory(dirName);
                Console.WriteLine("  file {0}", fileName);
                if (packagePart.ContentType.EndsWith("xml"))
                {
                    // open the XML from the Page Contents part
                    System.Xml.Linq.XDocument packagePartXML =
                        GetXDocFromPackagePart(packagePart);

                    // and save it to the file
                    // (the result is fine for me, but you might wanna use an XMLWriter for better/nicer formatting)
                    packagePartXML.Save(fileName);
                }
                else
                {
                    // just save the non XML as it is
                    FileStream newFileStrem =
                        new FileStream(fileName, FileMode.Create);
                    packagePart.GetStream().CopyTo(newFileStrem);
                }
            }
        } // static public void UnpackPackage(Package package, string targetDirectory)

        private static System.Xml.Linq.XDocument
        GetXDocFromPackagePart(PackagePart packagePart)
        {
            System.Xml.Linq.XDocument partXml = null;

            // read the XML document from the package part's stream
            Stream partStream = packagePart.GetStream();
            partXml = System.Xml.Linq.XDocument.Load(partStream);

            // Important: Close the stream or we will get an exception when writing the xml back to the package part.
            partStream.Close();
            return partXml;
        } // static private XDocument GetXDocFromPackagePart(PackagePart packagePart)

        // create the given directory / path
        static public int CreateDirectory(string path)
        {
            try
            {
                // Determine whether the directory exists.
                if (Directory.Exists(path))
                {
                    // TODO: use logger Console.WriteLine("That path \"{0}\" exists already.", path);
                    return 0;
                }
                // Try to create the directory.
                DirectoryInfo di = Directory.CreateDirectory(path);
                // TODO: use logger Console.WriteLine("The directory \"{0}\" was created successfully at {1}.", di.FullName, Directory.GetCreationTime(path));
                return 0;
            }
            catch (Exception e)
            {
                Console.WriteLine("CreateDirectory failed: {0}", e.ToString());
                return 1;
            }
            finally { }
        } // static public void CreateDirectory(string path)

    } // public class Program

} // namespace ClosedXML.WorkedExamples