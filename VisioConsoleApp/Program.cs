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

                //string dirPath = currentDirectory;

                DirectoryInfo currentDirectoryInfo = new DirectoryInfo(currentDirectory);

                // It is a best practice to get the file name string
                // using a FileInfo object, but it isn't necessary.
                FileInfo[] visioFileInfos = currentDirectoryInfo.GetFiles("*.vsdx");

                if (!visioFileInfos.Any())
                {
                    Console.WriteLine($"No *.vsdx files in folder: [{currentDirectoryInfo.FullName}]");
                }
                else
                {
                    //FileInfo firstVisioFileInfo = visioFileInfos[0];
                    //string firstVisioFileName = firstVisioFileInfo.Name;

                    ////// We're not going to do any more than open
                    ////// and read the list of parts in the package, although
                    ////// we can create a package or read/write what's inside.
                    ////using (Package fPackage = Package.Open(fName, FileMode.Open, FileAccess.Read))
                    ////{

                    ////    // The way to get a reference to a package part is
                    ////    // by using its URI. Thus, we're reading the URI
                    ////    // for each part in the package.
                    ////    PackagePartCollection fParts = fPackage.GetParts();
                    ////    foreach (PackagePart fPart in fParts)
                    ////    {
                    ////        Console.WriteLine("Package part: {0}", fPart.Uri);
                    ////    }
                    ////}

                    //var packageFileFullPath = GetPackagePath(firstVisioFileName, currentDirectory);

                    // Open the Visio file in a Package object.
                    //using Package visioPackage = OpenPackage(fName,  dirPath);

                    var packageFileFullPath = fileFullPath;

                    using Package visioFilePackage = Package.Open(packageFileFullPath, FileMode.OpenOrCreate, FileAccess.ReadWrite);

                    // Write the URI and content type of each package part to the console.
                    IteratePackageParts(visioFilePackage);

                    var customPropertiesPackagePartDocument = 
                        GetCustomPropertiesPackagePartDocument(visioFilePackage);

                    var willRecalcDocument = WillRecalcDocument(visioFilePackage);

                    // Get a reference to the Visio Document part contained in the file package.
                    PackagePart documentPackagePart = GetPackagePart(
                        visioFilePackage,
                        "http://schemas.microsoft.com/visio/2010/relationships/document");

                    // Get a reference to the collection of pages in the document, 
                    // and then to the first page in the document.
                    PackagePart pagesPackagePart = GetPackagePartFirstOrDefault(
                        visioFilePackage,
                        documentPackagePart,
                        "http://schemas.microsoft.com/visio/2010/relationships/pages");

                    PackagePart pagePackagePart = GetPackagePartFirstOrDefault(
                        visioFilePackage,
                        pagesPackagePart,
                        "http://schemas.microsoft.com/visio/2010/relationships/page");

                    using (var pagePackagePartStream = pagePackagePart.GetStream())
                    {
                        // Open the XML from the Page Contents part.
                        XDocument pageXDocument = GetXDocumentFromPartStream(pagePackagePartStream);

                        // Get all of the shapes from the page by getting
                        // all of the Shape elements from the pageXML document.
                        IEnumerable<XElement> shapeElements =
                            GetXElementsByLocalName(pageXDocument, "Shape")
                            .ToList();

                        //var shapeWithNameUAttributeElements =
                        //    shapeElements
                        //    .Where(e => e.Attribute("NameU") != null)
                        //    .ToList();

                        //// Select a Shape element from the shapes on the page by 
                        //// its name. You can modify this code to select elements
                        //// by other attributes and their values.
                        //XElement startEndShapeElement =
                        //    GetXElementByAttribute(shapeElements, "NameU", "Start/End");

                        // XName mainNamespace = "http://schemas.microsoft.com/office/visio/2012/main";
                        XNamespace mainNamespace = "http://schemas.microsoft.com/office/visio/2012/main";

                        // Edit Text Element
                        XName textXName = mainNamespace + "Text";

                        // Select the first Text Element
                        XElement firstTextElement =
                            shapeElements
                                .Descendants(textXName)
                                .FirstOrDefault();

                        // Get the Shape Parent Element
                        XElement startEndShapeElement = firstTextElement.Parent;

                        // Query the XML for the shape to get the Text element, and
                        // return the first Text element node.
                        IEnumerable<XElement> textElements =
                            startEndShapeElement
                            .Elements()
                            .Where(e => e.Name.LocalName == "Text")
                            .ToList();

                        var textElement = textElements.ElementAt(0);

                        // Change the shape text, leaving the <cp> element alone.
                        textElement.LastNode.ReplaceWith("Start process");

                        // Save the XML back to the Page Contents part.
                        SaveXDocumentToPackagePartStream(pageXDocument, pagePackagePartStream);

                        // Check update has occurred
                        pageXDocument = GetXDocumentFromPartStream(pagePackagePartStream);

                        // Add New Cell to Shape
                        // Insert a new Cell element in the Start/End shape that adds an arbitrary
                        // local ThemeIndex value. This code assumes that the shape does not 
                        // already have a local ThemeIndex cell.
                        startEndShapeElement.Add(new XElement("Cell",
                            new XAttribute("N", "ThemeIndex"),
                            new XAttribute("V", "25"),
                            new XProcessingInstruction("NewValue", "V")));

                        // Save the XML back to the Page Contents part.
                        SaveXDocumentToPackagePartStream(pageXDocument, pagePackagePartStream);

                        // Edit Shape
                        // Change the shape's horizontal position on the page 
                        // by getting a reference to the Cell element for the PinY 
                        // ShapeSheet cell and changing the value of its V attribute.
                        var pinYCellElement = 
                            GetXElementByAttribute(
                                startEndShapeElement.Elements(), 
                                "N", 
                                "PinY");

                        pinYCellElement.SetAttributeValue("V", "2");

                        //Save the XML back to the Page Contents part.
                        SaveXDocumentToPackagePartStream(pageXDocument, pagePackagePartStream);

                        //Add instructions to Visio to recalculate the entire document
                        //when it is next opened.
                        EnsureRecalcDocument(visioFilePackage);
                    }
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
            PackagePart sourcePackagePart,
            string relationship)
        {
            // This gets only the first PackagePart that shares the relationship
            // with the PackagePart passed in as an argument. You can modify the code
            // here to return a different PackageRelationship from the collection.
            var packageRelationship =
                sourcePackagePart
                    .GetRelationshipsByType(relationship)
                    .FirstOrDefault();

            PackagePart relatedPart = null;

            if (packageRelationship != null)
            {
                // Use the PackUriHelper class to determine the URI of PackagePart
                // that has the specified relationship to the PackagePart passed in
                // as an argument.
                Uri partUri = PackUriHelper.ResolvePartUri(
                    sourcePackagePart.Uri,
                    packageRelationship.TargetUri);

                relatedPart = filePackage.GetPart(partUri);
            }

            return relatedPart;
        }

        private static XDocument GetXDocumentFromPartStream(
            Stream packagePartStream)
        {
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

        private static IEnumerable<XElement> GetXElementsByLocalName(
            XDocument packagePartDocument,
            string elementLocalName)
        {
            // Construct a LINQ query that selects elements by their element type.
            IEnumerable<XElement> elements =
                packagePartDocument
                .Descendants()
                .Where(e => e.Name.LocalName == elementLocalName)
                .ToList();

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

        //private static void CopyStream(Stream input, Stream output)
        //{
        //    byte[] buffer = new byte[8192];
        //    int bytesRead;
        //    while ((bytesRead = input.Read(buffer, 0, buffer.Length)) > 0)
        //    {
        //        output.Write(buffer, 0, bytesRead);
        //    }
        //}

        private static bool WillRecalcDocument(Package filePackage)
        {
            var customPropertiesPackagePartDocument = 
                GetCustomPropertiesPackagePartDocument(filePackage);

            if (customPropertiesPackagePartDocument == null)
            {
                return false;
            }

            // Get all of the property elements from the document. 
            var propertyElements =
                GetXElementsByLocalName(customPropertiesPackagePartDocument, "property")
                .ToList();

            // Get the RecalcDocument property from the document if it exists already.
            var recalcDocumentPropertyElement =
                GetXElementByAttribute(
                    propertyElements,
                    "name",
                    "RecalcDocument");

            return recalcDocumentPropertyElement != null;
        }

        private static XDocument GetCustomPropertiesPackagePartDocument(
            Package filePackage)
        {
            PackagePart customPropertiesPackagePart = 
                GetPackagePart(
                    filePackage,
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/" + "custom-properties");

            using (var customPropertiesPackagePartStream = customPropertiesPackagePart.GetStream())
            {
                var customPropertiesPackagePartDocument = 
                    GetXDocumentFromPartStream(customPropertiesPackagePartStream);

                return customPropertiesPackagePartDocument;
            }
        }

        private static void EnsureRecalcDocument(
            Package filePackage)
        {
            PackagePart customPropertiesPackagePart = GetPackagePart(filePackage,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/" +
                "custom-properties");

            using (var customPropertiesPackagePartStream = customPropertiesPackagePart.GetStream())
            {
                var customPropertiesPackagePartDocument =
                    GetXDocumentFromPartStream(customPropertiesPackagePartStream);

                // Check to see whether document recalculation has already been 
                // set for this document. If it hasn't, use the integer
                // value returned by CheckForRecalc as the property ID.
                int pidValue = GetPropertyIdIntegerForRecalcDocument(customPropertiesPackagePartDocument);

                if (pidValue > -1)
                {
                    // RecalcDocument is not present
                    // Add it with the computed pid value
                    XElement customPartRoot = 
                        customPropertiesPackagePartDocument
                            .Elements()
                            .ElementAt(0);

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
                SaveXDocumentToPackagePartStream(
                    customPropertiesPackagePartDocument, 
                    customPropertiesPackagePartStream);
            }
        }

        private static void SaveXDocumentToPackagePartStream(
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

        private static int GetPropertyIdIntegerForRecalcDocument(XDocument customPropertiesDocument)
        {
            // Set the inital pidValue to -1, which is not an allowed value.
            // The calling code tests to see whether the pidValue is 
            // greater than -1, where -1 means that the RecalcDocument property is already present.
            int pidValue = -1;

            // Get all of the property elements from the document. 
            var propertyElements = 
                GetXElementsByLocalName(customPropertiesDocument, "property")
                .ToList();

            // Get the RecalcDocument property from the document, if it exists already.
            var recalcDocumentPropertyElement = 
                GetXElementByAttribute(
                    propertyElements, 
                    "name", 
                    "RecalcDocument");

            // If there is already a RecalcDocument instruction in the 
            // Custom File Properties part, then we don't need to add another one. 
            // Otherwise, we need to return a unique pid value.
            if (recalcDocumentPropertyElement != null)
            {
                // Client should not add RecalcDocument element
                return -1;
            }

            // Get all of the pid values of the property elements and then
            // convert the IEnumerable object into an array.
            //var propertyIds =
            //    from prop in propertyElements
            //    where prop != null
            //    where prop.Name != null
            //    where prop.Name.LocalName == "property"
            //    select prop.Attribute("pid")?.Value;

            //string[] propertyIdArray = propertyIds.ToArray();

            var propertyIdStrings =
                propertyElements
                .Where(i => i != null)
                .Where(i => i.Name != null)
                .Where(i => i.Name.LocalName == "property")
                .Select(i => i.Attribute("pid")?.Value)
                .Where(pid => !string.IsNullOrWhiteSpace(pid))
                .ToList();

            var propertyIdIntegers =
                propertyIdStrings
                .Select(s => int.TryParse(s, out int parsedResult) ? parsedResult : 0)
                .OrderBy(i => i)
                .ToList();

            var maximumExistingId = propertyIdIntegers.Max();

            // Increment this id value until a unique value is found.
            // This starts at 2, because 0 and 1 are not valid pid values.
            var minimumId = 2;

            pidValue = Math.Max(minimumId, maximumExistingId + 1);

            //while (pidValue == -1)
            //{
            //    if (propertyIdArray.Contains(id.ToString()))
            //    {
            //        id++;
            //    }
            //    else
            //    {
            //        pidValue = id;
            //    }
            //}

            return pidValue;
        }
    }
}
