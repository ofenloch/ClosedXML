
using System;
using System.IO;
using System.IO.Packaging;

namespace TemplateForGeWi
{
    class Utilities
    {
        public static string BaseCreatedDirectory
        {
            get
            {
                //var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Created");
                var path = PathCombine(Directory.GetCurrentDirectory(), "output/Created");
                if (!Directory.Exists(path)) Directory.CreateDirectory(path);
                return path;
            }
        } // public static string BaseCreatedDirectory

        public static string BaseModifiedDirectory
        {
            get
            {
                //var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Modified");
                var path = PathCombine(Directory.GetCurrentDirectory(), "output/Modified");
                if (!Directory.Exists(path)) Directory.CreateDirectory(path);
                return path;
            }
        } // public static string BaseModifiedDirectory

        /// <summary>
        /// make sure the returned string has only directory separators valid on the current platform
        /// </summary>
        /// <remarks>
        /// <para>As a string representing a path segment</para>
        /// </remarks>
        public static string PathAdjustSeparators(string path)
        {
            string newPath = path.Replace('/', Path.DirectorySeparatorChar);
            newPath = newPath.Replace('\\', Path.DirectorySeparatorChar);
            return newPath;
        }

        /// <summary>
        /// replacement for Path.Combine that should work with mixed path separators
        /// </summary>
        /// <remarks>
        /// <para>As a string representing a path segment</para>
        /// <para>As a string representing a path segment</para>
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
                string? dirName = Path.GetDirectoryName(fileName);
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
            System.Xml.Linq.XDocument? partXml = null;

            // read the XML document from the package part's stream
            Stream partStream = packagePart.GetStream();
            partXml = System.Xml.Linq.XDocument.Load(partStream);

            // Important: Close the stream or we will get an exception when writing the xml back to the package part.
            partStream.Close();
            return partXml;
        } // static private XDocument GetXDocFromPackagePart(PackagePart packagePart)

        // create the given directory / path
        static public int CreateDirectory(string? path)
        {
            if (String.IsNullOrEmpty(path))
            {
                return 200;
            }
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

    } // class Utilities

} // namespace TemplateForGeWi