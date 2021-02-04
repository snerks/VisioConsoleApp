using System;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using System.Text;
using System.Collections.Generic;

namespace VisioConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            var currentDirectory = Environment.CurrentDirectory;

            var fileName = "tf66758849.vsdx";

            var fileFullPath = System.IO.Path.Combine(currentDirectory, fileName);

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

                    // Open the Visio file in a Package object.
                    using (Package visioPackage = OpenPackage(
                        fName,
                        dirPath))
                    {
                        // Write the URI and content type of each package part to the console.
                        IteratePackageParts(visioPackage);

                        // Get a reference to the Visio Document part contained in the file package.
                        PackagePart documentPart = GetPackagePart(
                            visioPackage,
                            "http://schemas.microsoft.com/visio/2010/relationships/document");

                        // Get a reference to the collection of pages in the document, 
                        // and then to the first page in the document.
                        PackagePart pagesPart = GetPackagePart(
                            visioPackage, 
                            documentPart,
                            "http://schemas.microsoft.com/visio/2010/relationships/pages");

                        PackagePart pagePart = GetPackagePart(
                            visioPackage, 
                            pagesPart,
                            "http://schemas.microsoft.com/visio/2010/relationships/page");

                        // Open the XML from the Page Contents part.
                        XDocument pageXML = GetXMLFromPart(pagePart);

                        // Get all of the shapes from the page by getting
                        // all of the Shape elements from the pageXML document.
                        IEnumerable<XElement> shapeElements = 
                            GetXElementsByName(pageXML, "Shape")
                            .ToList();

                        // Select a Shape element from the shapes on the page by 
                        // its name. You can modify this code to select elements
                        // by other attributes and their values.
                        //XElement startEndShapeXML =
                        //    GetXElementByAttribute(shapesXML, "NameU", "Start/End");

                        XElement startEndShapeElement =
                            shapeElements.FirstOrDefault();

                        // Query the XML for the shape to get the Text element, and
                        // return the first Text element node.
                        IEnumerable<XElement> textElements = from element in startEndShapeElement.Elements()
                                                             where element.Name.LocalName == "Text"
                                                             select element;

                        XElement textElement = textElements.ElementAt(0);

                        // Change the shape text, leaving the <cp> element alone.
                        textElement.LastNode.ReplaceWith("Start process");

                        // Save the XML back to the Page Contents part.
                        SaveXDocumentToPart(pagePart, pageXML);

                        // Insert a new Cell element in the Start/End shape that adds an arbitrary
                        // local ThemeIndex value. This code assumes that the shape does not 
                        // already have a local ThemeIndex cell.
                        startEndShapeElement.Add(new XElement("Cell",
                            new XAttribute("N", "ThemeIndex"),
                            new XAttribute("V", "25"),
                            new XProcessingInstruction("NewValue", "V")));

                        // Save the XML back to the Page Contents part.
                        SaveXDocumentToPart(pagePart, pageXML);

                        //// Change the shape's horizontal position on the page 
                        //// by getting a reference to the Cell element for the PinY 
                        //// ShapeSheet cell and changing the value of its V attribute.
                        //XElement pinYCellXML = GetXElementByAttribute(
                        //    startEndShapeElement.Elements(), "N", "PinY");

                        //pinYCellXML.SetAttributeValue("V", "2");

                        // Add instructions to Visio to recalculate the entire document
                        // when it is next opened.
                        RecalcDocument(visioPackage);
                        
                        // Save the XML back to the Page Contents part.
                        SaveXDocumentToPart(pagePart, pageXML);
                    }
                }
            }
            catch (Exception err)
            {
                Console.WriteLine("Error: {0}", err.Message);
            }
            finally
            {
                Console.Write("\nPress any key to continue ...");
                Console.ReadKey();
            }

            Console.WriteLine(currentDirectory);
            Console.ReadKey();
        }

        private static Package OpenPackage(
            string fileName,
            string folderFullPath)
                        //Environment.SpecialFolder folder)
        {
            Package visioPackage = null;
            // Get a reference to the location 
            // where the Visio file is stored.
            //string directoryPath = System.Environment.GetFolderPath(folder);
            string directoryPath = folderFullPath;

            DirectoryInfo dirInfo = new DirectoryInfo(directoryPath);

            // Get the Visio file from the location.
            FileInfo[] fileInfos = dirInfo.GetFiles(fileName);
            
            if (fileInfos.Count() > 0)
            {
                FileInfo fileInfo = fileInfos[0];
                string filePathName = fileInfo.FullName;
                // Open the Visio file as a package with
                // read/write file access.
                visioPackage = Package.Open(
                    filePathName,
                    FileMode.Open,
                    FileAccess.ReadWrite);
            }
            
            // Return the Visio file as a package.
            return visioPackage;
        }

        private static void IteratePackageParts(Package filePackage)
        {
            // Get all of the package parts contained in the package
            // and then write the URI and content type of each one to the console.
            PackagePartCollection packageParts = filePackage.GetParts();

            foreach (PackagePart part in packageParts)
            {
                Console.WriteLine("Package part URI: {0}", part.Uri);
                Console.WriteLine("Content type: {0}", part.ContentType.ToString());
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

        private static PackagePart GetPackagePart(
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

        private static XDocument GetXMLFromPart(PackagePart packagePart)
        {
            XDocument partXml = null;
            // Open the packagePart as a stream and then 
            // open the stream in an XDocument object.
            Stream partStream = packagePart.GetStream();
            partXml = XDocument.Load(partStream);
            return partXml;
        }

        private static IEnumerable<XElement> GetXElementsByName(
            XDocument packagePart, 
            string elementType)
        {
            // Construct a LINQ query that selects elements by their element type.
            IEnumerable<XElement> elements =
                from element in packagePart.Descendants()
                where element.Name.LocalName == elementType
                select element;

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
            IEnumerable<XElement> selectedElements =
                from el in elements
                where el.Attribute(attributeName).Value == attributeValue
                select el;

            // If there aren't any elements of the specified type
            // with the specified attribute value in the document,
            // return null to the calling code.
            return selectedElements.DefaultIfEmpty(null).FirstOrDefault();
        }

        private static void SaveXDocumentToPart(
            PackagePart packagePart,
            XDocument partXML)
        {

            // Create a new XmlWriterSettings object to 
            // define the characteristics for the XmlWriter
            XmlWriterSettings partWriterSettings = new XmlWriterSettings();
            partWriterSettings.Encoding = Encoding.UTF8;
            // Create a new XmlWriter and then write the XML
            // back to the document part.
            XmlWriter partWriter = XmlWriter.Create(packagePart.GetStream(),
                partWriterSettings);
            partXML.WriteTo(partWriter);
            // Flush and close the XmlWriter.
            partWriter.Flush();
            partWriter.Close();
        }

        private static void RecalcDocument(Package filePackage)
        {
            // Get the Custom File Properties part from the package and
            // and then extract the XML from it.
            PackagePart customPart = GetPackagePart(filePackage,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/" +
                "custom-properties");
            XDocument customPartXML = GetXMLFromPart(customPart);
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
                customPartRoot.Add(
                    new XElement(customPropsSchemaNS + "property",
                        new XAttribute("pid", pidValue.ToString()),
                        new XAttribute("name", "RecalcDocument"),
                        new XAttribute("fmtid",
                            "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"),
                        new XElement(customVTypesNS + "bool", "true")
                    ));
            }
            // Save the Custom Properties package part back to the package.
            SaveXDocumentToPart(customPart, customPartXML);
        }

        private static int CheckForRecalc(XDocument customPropsXDoc)
        {
            // Set the inital pidValue to -1, which is not an allowed value.
            // The calling code tests to see whether the pidValue is 
            // greater than -1.
            int pidValue = -1;
            // Get all of the property elements from the document. 
            IEnumerable<XElement> props = GetXElementsByName(
                customPropsXDoc, "property");
            // Get the RecalcDocument property from the document if it exists already.
            XElement recalcProp = GetXElementByAttribute(props,
                "name", "RecalcDocument");
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
                    from prop in props
                    where prop.Name.LocalName == "property"
                    select prop.Attribute("pid").Value;
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
