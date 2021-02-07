using System;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using System.Text;
using System.Collections.Generic;
using System.IO.Compression;

namespace VisioConsoleApp
{
    internal static class Program
    {
        private static void Main()
        {
            var currentDirectory = Environment.CurrentDirectory;

            const string fileName = "tf66758849.vsdx";

            var fileFullPath = Path.Combine(currentDirectory, fileName);

            try
            {
                Console.WriteLine("Opening the VSDX file ...");

                // Need to get the folder path for the Desktop
                // where the file is saved.
                //string dirPath = System.Environment.GetFolderPath(
                //    System.Environment.SpecialFolder.Desktop);

                string dirPath = currentDirectory;

                DirectoryInfo myDir = new DirectoryInfo(dirPath);

                // It is a best practice to get the file name string
                // using a FileInfo object, but it isn't necessary.
                FileInfo[] fInfos = myDir.GetFiles("*.vsdx");

                if (!fInfos.Any())
                {
                    Console.WriteLine($"No *.vsdx files in folder: [{myDir.FullName}]");
                }
                else
                {
                    FileInfo fi = fInfos[0];
                    string fName = fi.Name;

                    //// We're not going to do any more than open
                    //// and read the list of parts in the package, although
                    //// we can create a package or read/write what's inside.
                    //using (Package fPackage = Package.Open(fName, FileMode.Open, FileAccess.Read))
                    //{

                    //    // The way to get a reference to a package part is
                    //    // by using its URI. Thus, we're reading the URI
                    //    // for each part in the package.
                    //    PackagePartCollection fParts = fPackage.GetParts();
                    //    foreach (PackagePart fPart in fParts)
                    //    {
                    //        Console.WriteLine("Package part: {0}", fPart.Uri);
                    //    }
                    //}

                    var packageFileFullPath = GetPackagePath(fName, dirPath);

                    // Open the Visio file in a Package object.
                    //using Package visioPackage = OpenPackage(fName,  dirPath);
                    using Package visioPackage = Package.Open(packageFileFullPath, FileMode.OpenOrCreate, FileAccess.ReadWrite);

                    // Write the URI and content type of each package part to the console.
                    IteratePackageParts(visioPackage);

                    // Get a reference to the Visio Document part contained in the file package.
                    PackagePart documentPart = GetPackagePart(
                        visioPackage,
                        "http://schemas.microsoft.com/visio/2010/relationships/document");

                    // Get a reference to the collection of pages in the document, 
                    // and then to the first page in the document.
                    PackagePart pagesPart = GetPackagePartFirstOrDefault(
                        visioPackage,
                        documentPart,
                        "http://schemas.microsoft.com/visio/2010/relationships/pages");

                    PackagePart pagePart = GetPackagePartFirstOrDefault(
                        visioPackage,
                        pagesPart,
                        "http://schemas.microsoft.com/visio/2010/relationships/page");

                    using (var packageStream = pagePart.GetStream())
                    {
                        //do stuff with stream - not necessary to reproduce bug

                        // Open the XML from the Page Contents part.
                        XDocument pageXDocument = GetXDocumentFromPartStream(packageStream);

                        // Get all of the shapes from the page by getting
                        // all of the Shape elements from the pageXML document.
                        IEnumerable<XElement> shapeElements =
                            GetXElementsByName(pageXDocument, "Shape")
                            .ToList();

                        // Select a Shape element from the shapes on the page by 
                        // its name. You can modify this code to select elements
                        // by other attributes and their values.
                        //XElement startEndShapeXML =
                        //    GetXElementByAttribute(shapesXML, "NameU", "Start/End");

                        // XName mainNamespace = "http://schemas.microsoft.com/office/visio/2012/main";
                        XNamespace mainNamespace = "http://schemas.microsoft.com/office/visio/2012/main";

                        XElement firstTextElement =
                            shapeElements
                                .Descendants(mainNamespace + "Text")
                                .FirstOrDefault();

                        XElement startEndShapeElement = firstTextElement.Parent;

                        // Query the XML for the shape to get the Text element, and
                        // return the first Text element node.
                        IEnumerable<XElement> textElements = from element in startEndShapeElement.Elements()
                                                             where element.Name.LocalName == "Text"
                                                             select element;

                        XElement textElement = textElements.ElementAt(0);

                        // Change the shape text, leaving the <cp> element alone.
                        textElement.LastNode.ReplaceWith("Start process");

                        // Save the XML back to the Page Contents part.
                        //SaveXDocumentToPart(packageStream, pagePart, pageXDocument);
                        SaveXDocumentToPart(pageXDocument, packageStream);

                        pageXDocument = GetXDocumentFromPartStream(packageStream);

                        // Insert a new Cell element in the Start/End shape that adds an arbitrary
                        // local ThemeIndex value. This code assumes that the shape does not 
                        // already have a local ThemeIndex cell.
                        startEndShapeElement.Add(new XElement("Cell",
                            new XAttribute("N", "ThemeIndex"),
                            new XAttribute("V", "25"),
                            new XProcessingInstruction("NewValue", "V")));

                        // Save the XML back to the Page Contents part.
                        //SaveXDocumentToPart(packageStream, pagePart, pageXDocument);
                        SaveXDocumentToPart(pageXDocument, packageStream);

                        // Change the shape's horizontal position on the page 
                        // by getting a reference to the Cell element for the PinY 
                        // ShapeSheet cell and changing the value of its V attribute.
                        XElement pinYCellXML = GetXElementByAttribute(
                            startEndShapeElement.Elements(), "N", "PinY");

                        pinYCellXML.SetAttributeValue("V", "2");

                        //Add instructions to Visio to recalculate the entire document
                        //when it is next opened.
                        RecalcDocument(visioPackage, packageStream);

                        //Save the XML back to the Page Contents part.
                        //SaveXDocumentToPart(packageStream, pagePart, pageXDocument);
                        SaveXDocumentToPart(pageXDocument, packageStream);
                    }

                    //using (var packageStream = pagePart.GetStream(FileMode.Open))
                    //{
                    //    //do stuff with stream - not necessary to reproduce bug
                    //    // Save the XML back to the Page Contents part.
                    //    SaveXDocumentToPart(packageStream, pagePart, pageXML);
                    //}

                    //// Save the XML back to the Page Contents part.
                    //SaveXDocumentToPart(visioPackage, pagePart, pageXML);

                    //// Insert a new Cell element in the Start/End shape that adds an arbitrary
                    //// local ThemeIndex value. This code assumes that the shape does not 
                    //// already have a local ThemeIndex cell.
                    //startEndShapeElement.Add(new XElement("Cell",
                    //    new XAttribute("N", "ThemeIndex"),
                    //    new XAttribute("V", "25"),
                    //    new XProcessingInstruction("NewValue", "V")));

                    //// Save the XML back to the Page Contents part.
                    ////SaveXDocumentToPart(visioPackage, pagePart, pageXML);

                    //// Change the shape's horizontal position on the page 
                    //// by getting a reference to the Cell element for the PinY 
                    //// ShapeSheet cell and changing the value of its V attribute.
                    //XElement pinYCellXML = GetXElementByAttribute(
                    //    startEndShapeElement.Elements(), "N", "PinY");

                    //pinYCellXML.SetAttributeValue("V", "2");

                    ////Add instructions to Visio to recalculate the entire document
                    ////when it is next opened.
                    //RecalcDocument(visioPackage);

                    ////Save the XML back to the Page Contents part.
                    ////SaveXDocumentToPart(visioPackage, pagePart, pageXML);
                }
            }
            catch (Exception err)
            {
                Console.WriteLine("Error: {0}", err.Message);
            }
            finally
            {
                Console.WriteLine(currentDirectory);
                Console.Write("\nPress any key to continue ...");
                Console.ReadKey();
            }

            //Console.ReadKey();
        }

        //private static Package OpenPackage(
        //    string fileName,
        //    string folderFullPath)
        //                //Environment.SpecialFolder folder)
        //{
        //    Package visioPackage = null;
        //    // Get a reference to the location 
        //    // where the Visio file is stored.
        //    //string directoryPath = System.Environment.GetFolderPath(folder);
        //    string directoryPath = folderFullPath;

        //    DirectoryInfo dirInfo = new DirectoryInfo(directoryPath);

        //    // Get the Visio file from the location.
        //    FileInfo[] fileInfos = dirInfo.GetFiles(fileName);

        //    if (fileInfos.Length > 0)
        //    {
        //        FileInfo fileInfo = fileInfos[0];
        //        string filePathName = fileInfo.FullName;

        //        // Open the Visio file as a package with
        //        // read/write file access.
        //        visioPackage = Package.Open(
        //            filePathName,
        //            FileMode.Open,
        //            FileAccess.ReadWrite);
        //        //FileAccess.Read);
        //    }

        //    // Return the Visio file as a package.
        //    return visioPackage;
        //}

        private static string GetPackagePath(
            string fileName,
            string folderFullPath)
            //Environment.SpecialFolder folder)
        {
            //Package visioPackage = null;
            // Get a reference to the location 
            // where the Visio file is stored.
            //string directoryPath = System.Environment.GetFolderPath(folder);
            string directoryPath = folderFullPath;

            DirectoryInfo dirInfo = new DirectoryInfo(directoryPath);

            string filePathName = null;

            // Get the Visio file from the location.
            FileInfo[] fileInfos = dirInfo.GetFiles(fileName);

            if (fileInfos.Length > 0)
            {
                FileInfo fileInfo = fileInfos[0];
                filePathName = fileInfo.FullName;

                // Open the Visio file as a package with
                // read/write file access.
                //visioPackage = Package.Open(
                //    filePathName,
                //    FileMode.Open,
                //    FileAccess.ReadWrite);
                //FileAccess.Read);
            }

            // Return the Visio file as a package.
            //return visioPackage;

            return filePathName;
        }

        private static void IteratePackageParts(Package filePackage)
        {
            // Get all of the package parts contained in the package
            // and then write the URI and content type of each one to the console.
            PackagePartCollection packageParts = filePackage.GetParts();

            foreach (PackagePart part in packageParts)
            {
                Console.WriteLine("Package part URI: {0}", part.Uri);
                Console.WriteLine("Content type: {0}", part.ContentType);
            }
        }

        private static PackagePart GetPackagePart(
            Package filePackage,
            string relationship)
        {
            // Use the namespace that describes the relationship 
            // to get the relationship.
            PackageRelationship packageRel =
                filePackage
                    .GetRelationshipsByType(relationship)
                    .FirstOrDefault();

            PackagePart part = null;

            // If the Visio file package contains this type of relationship with 
            // one of its parts, return that part.
            if (packageRel != null)
            {
                // Clean up the URI using a helper class and then get the part.
                Uri docUri = PackUriHelper.ResolvePartUri(
                    new Uri("/", UriKind.Relative),
                    packageRel.TargetUri);

                part = filePackage.GetPart(docUri);
            }

            return part;
        }

        private static PackagePart GetPackagePartFirstOrDefault(
            Package filePackage,
            PackagePart sourcePart,
            string relationship)
        {
            // This gets only the first PackagePart that shares the relationship
            // with the PackagePart passed in as an argument. You can modify the code
            // here to return a different PackageRelationship from the collection.
            PackageRelationship packageRel =
                sourcePart.GetRelationshipsByType(relationship).FirstOrDefault();

            PackagePart relatedPart = null;

            if (packageRel != null)
            {
                // Use the PackUriHelper class to determine the URI of PackagePart
                // that has the specified relationship to the PackagePart passed in
                // as an argument.
                Uri partUri = PackUriHelper.ResolvePartUri(
                    sourcePart.Uri,
                    packageRel.TargetUri);

                relatedPart = filePackage.GetPart(partUri);
            }
            return relatedPart;
        }

        private static XDocument GetXDocumentFromPartStream(
            //PackagePart packagePart,
            Stream packagePartStream)
        {
            // Open the packagePart as a stream and then 
            // open the stream in an XDocument object.
            //Stream partStream = packagePart.GetStream();

            packagePartStream.Position = 0;

            var buffer = new byte[packagePartStream.Length];

            packagePartStream.Read(buffer, 0, (int)packagePartStream.Length);

            string xml = System.Text.Encoding.UTF8.GetString(buffer);

            var preamble = Encoding.UTF8.GetPreamble();

            if (buffer[0] == preamble[0])
            {
                var stop = true;
            }

            //string byteOrderMarkUtf8 = Encoding.UTF8.GetString(Encoding.UTF8.GetPreamble());
            //if (xml.StartsWith(byteOrderMarkUtf8, StringComparison.Ordinal))
            //{
            //    // xml = _byteOrderMarkUtf8 + xml;
            //    xml = xml.Remove(0, byteOrderMarkUtf8.Length);
            //}

            packagePartStream.Position = 0;

            try
            {
                var partXDocument = XDocument.Load(packagePartStream);
                //var partXDocument = XDocument.Parse(xml);

                return partXDocument;
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        private static IEnumerable<XElement> GetXElementsByName(
            XDocument packagePartDocument,
            string elementLocalName)
        {
            // Construct a LINQ query that selects elements by their element type.
            IEnumerable<XElement> elements =
                (
                    from element in packagePartDocument.Descendants()
                    where element.Name.LocalName == elementLocalName
                    select element
                ).ToList();

            // Return the selected elements to the calling code.
            return elements.DefaultIfEmpty(null);
        }

        private static XElement GetXElementByAttribute(
            IEnumerable<XElement> elements,
            string attributeName,
            string attributeValue)
        {
            // Construct a LINQ query that selects elements from a group
            // of elements by the value of a specific attribute.
            //IEnumerable<XElement> selectedElements =
            //    from el in elements
            //    where el.Attribute(attributeName)?.Value == attributeValue
            //    select el;

            var selectedElements =
                elements
                    .Where(el => el != null)
                    .Where(el => el.Attribute(attributeName) != null)
                    .Where(el => el.Attribute(attributeName).Value == attributeValue)
                    .ToList();

            // If there aren't any elements of the specified type
            // with the specified attribute value in the document,
            // return null to the calling code.
            return selectedElements.DefaultIfEmpty(null).FirstOrDefault();
        }

        //private static void SaveXDocumentToPart(
        //    Stream packageStream,
        //    PackagePart packagePart,
        //    XDocument partXML)
        //{
        //    // Create a new XmlWriterSettings object to 
        //    // define the characteristics for the XmlWriter
        //    XmlWriterSettings partWriterSettings = new XmlWriterSettings
        //    {
        //        Encoding = Encoding.UTF8
        //    };

        //    try
        //    {
        //        //using (Stream file = File.OpenRead(filename))
        //        //using (Stream gzip = new GZipStream(file, CompressionMode.Decompress))
        //        using Stream memoryStream = new MemoryStream();

        //        //var packageStream = packagePart.GetStream();
        //        //using (var packageStream = packagePart.GetStream(FileMode.Open))
        //        //{
        //            //do stuff with stream - not necessary to reproduce bug
        //            CopyStream(packageStream, memoryStream);
        //            //return memoryStream.ToArray();

        //            // Create a new XmlWriter and then write the XML
        //            // back to the document part.
        //            XmlWriter partWriter =
        //                XmlWriter.Create(
        //                    //packagePart.GetStream(),
        //                    memoryStream,
        //                    partWriterSettings);

        //            partXML.WriteTo(partWriter);

        //            // Flush and close the XmlWriter.
        //            partWriter.Flush();
        //            partWriter.Close();

        //            CopyStream(memoryStream, packageStream);
        //        //}
        //    }
        //    catch (Exception ex)
        //    {
        //        throw;
        //    }
        //}

        //private static void SaveXDocumentToPart(
        //    Stream packageStream,
        //    PackagePart packagePart,
        //    XDocument partXML)
        //{
        //    // Create a new XmlWriterSettings object to 
        //    // define the characteristics for the XmlWriter
        //    XmlWriterSettings partWriterSettings = new XmlWriterSettings
        //    {
        //        Encoding = Encoding.UTF8
        //    };

        //    try
        //    {
        //        ////using (Stream file = File.OpenRead(filename))
        //        ////using (Stream gzip = new GZipStream(file, CompressionMode.Decompress))
        //        //using Stream memoryStream = new MemoryStream();

        //        ////var packageStream = packagePart.GetStream();
        //        ////using (var packageStream = packagePart.GetStream(FileMode.Open))
        //        ////{
        //        ////do stuff with stream - not necessary to reproduce bug
        //        //CopyStream(packageStream, memoryStream);
        //        ////return memoryStream.ToArray();

        //        // Create a new XmlWriter and then write the XML
        //        // back to the document part.
        //        XmlWriter partWriter =
        //            XmlWriter.Create(
        //                //packagePart.GetStream(),
        //                packageStream,
        //                partWriterSettings);

        //        partXML.WriteTo(partWriter);

        //        // Flush and close the XmlWriter.
        //        partWriter.Flush();
        //        partWriter.Close();

        //        //CopyStream(memoryStream, packageStream);
        //        //}
        //    }
        //    catch (Exception ex)
        //    {
        //        throw;
        //    }
        //}

        private static void CopyStream(Stream input, Stream output)
        {
            byte[] buffer = new byte[8192];
            int bytesRead;
            while ((bytesRead = input.Read(buffer, 0, buffer.Length)) > 0)
            {
                output.Write(buffer, 0, bytesRead);
            }
        }

        private static void RecalcDocument(
            Package filePackage,
            Stream packageStream)
        {
            // Get the Custom File Properties part from the package and
            // and then extract the XML from it.
            PackagePart customPart = GetPackagePart(filePackage,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/" +
                "custom-properties");

            //XDocument customPartXML = GetXMLFromPart(customPart);
            XDocument customPartXML = GetXDocumentFromPartStream(packageStream);

            // Check to see whether document recalculation has already been 
            // set for this document. If it hasn't, use the integer
            // value returned by CheckForRecalc as the property ID.
            int pidValue = CheckForRecalc(customPartXML);
            if (pidValue > -1)
            {
                XElement customPartRoot = customPartXML.Elements().ElementAt(0);
                // Two XML namespaces are needed to add XML data to this 
                // document. Here, we're using the GetNamespaceOfPrefix and 
                // GetDefaultNamespace methods to get the namespaces that 
                // we need. You can specify the exact strings for the 
                // namespaces, but that is not recommended.
                XNamespace customVTypesNS = customPartRoot.GetNamespaceOfPrefix("vt");
                XNamespace customPropsSchemaNS = customPartRoot.GetDefaultNamespace();

                // Construct the XML for the new property in the XDocument.Add method.
                // This ensures that the XNamespace objects will resolve properly, 
                // apply the correct prefix, and will not default to an empty namespace.

                if (customVTypesNS != null)
                {
                    customPartRoot.Add(
                        new XElement(customPropsSchemaNS + "property",
                            new XAttribute("pid", pidValue.ToString()),
                            new XAttribute("name", "RecalcDocument"),
                            new XAttribute("fmtid",
                                "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"),
                            new XElement(customVTypesNS + "bool", "true")
                        ));
                }
            }

            // Save the Custom Properties package part back to the package.
            SaveXDocumentToPart(customPartXML, packageStream);
        }

        private static void SaveXDocumentToPart(
                // PackagePart packagePart,
                XDocument packagePartXDocument,
                Stream packagePartStream
            )
        {
            // Create a new XmlWriterSettings object to 
            // define the characteristics for the XmlWriter
            XmlWriterSettings partWriterSettings = new XmlWriterSettings();
            partWriterSettings.Encoding = Encoding.UTF8;
            partWriterSettings.Indent = true;

            packagePartStream.Position = 0;

            // Create a new XmlWriter and then write the XML
            // back to the document part.
            XmlWriter partWriter = 
                XmlWriter.Create(
                    //packagePart.GetStream(),
                    packagePartStream,
                    partWriterSettings);

            packagePartXDocument.WriteTo(partWriter);

            // Flush and close the XmlWriter.
            partWriter.Flush();
            partWriter.Close();
        }

        private static int CheckForRecalc(XDocument customPropsXDoc)
        {
            // Set the inital pidValue to -1, which is not an allowed value.
            // The calling code tests to see whether the pidValue is 
            // greater than -1.
            int pidValue = -1;

            // Get all of the property elements from the document. 
            var propElements = 
                GetXElementsByName(customPropsXDoc, "property")
                .ToList();

            // Get the RecalcDocument property from the document if it exists already.
            XElement recalcProp = 
                GetXElementByAttribute(propElements, "name", "RecalcDocument");

            // If there is already a RecalcDocument instruction in the 
            // Custom File Properties part, then we don't need to add another one. 
            // Otherwise, we need to create a unique pid value.
            if (recalcProp != null)
            {
                return pidValue;
            }
            else
            {
                // Get all of the pid values of the property elements and then
                // convert the IEnumerable object into an array.
                IEnumerable<string> propIDs =
                    from prop in propElements
                    where prop != null
                    where prop.Name != null
                    where prop.Name.LocalName == "property"
                    select prop.Attribute("pid")?.Value;

                string[] propIDArray = propIDs.ToArray();
                // Increment this id value until a unique value is found.
                // This starts at 2, because 0 and 1 are not valid pid values.
                int id = 2;
                while (pidValue == -1)
                {
                    if (propIDArray.Contains(id.ToString()))
                    {
                        id++;
                    }
                    else
                    {
                        pidValue = id;
                    }
                }
            }
            return pidValue;
        }
    }
}
