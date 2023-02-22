using System.Xml;
using System.Xml.Linq;
using System.IO.Packaging;
using System.Text;
using System.Diagnostics;


// recreation of the methods from https://learn.microsoft.com/en-us/office/client-developer/visio/how-to-manipulate-the-visio-file-format-programmatically

// this file contains the examples of the methods for editing and saving the file

namespace Helpers
{
    public class VisioHelper_NextSteps
    {
        Package? _visioPackage;
        IEnumerable<string>? _pagesNames;

        public void OpenVisioFile(string fileName)
        {

            using (_visioPackage = OpenPackage(fileName))
            {
                IEnumerable<PackagePart> pages = GetPages();

                PackagePart pagePart = pages.First();

                XDocument pageXML = GetXMLFromPart(pagePart);

                IEnumerable<XElement> shapesXML = GetShapesXml(pageXML);

                IEnumerable<string> texts = GetPrefixedTextsFromShapes(shapesXML, "Note");


                XElement startEndShapeXML = GetXElementByAttribute(shapesXML, "NameU", "Start/End");



                IEnumerable<XElement> textElements = from element in startEndShapeXML.Elements()
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
                startEndShapeXML.Add(new XElement("Cell",
                    new XAttribute("N", "ThemeIndex"),
                    new XAttribute("V", "25"),
                    new XProcessingInstruction("NewValue", "V")));
                // Save the XML back to the Page Contents part.
                SaveXDocumentToPart(pagePart, pageXML);


                // Change the shape's horizontal position on the page 
                // by getting a reference to the Cell element for the PinY 
                // ShapeSheet cell and changing the value of its V attribute.
                XElement pinYCellXML = GetXElementByAttribute(
                    startEndShapeXML.Elements(), "N", "PinY");
                pinYCellXML.SetAttributeValue("V", "2");
                // Add instructions to Visio to recalculate the entire document
                // when it is next opened.
                RecalcDocument();
                // Save the XML back to the Page Contents part.
                SaveXDocumentToPart(pagePart, pageXML);



                // Create a new Ribbon Extensibility part and add it to the file.
                XDocument customUIXML = CreateCustomUI();
                CreateNewPackagePart(_visioPackage, customUIXML,
                    new Uri("/customUI/customUI1.xml", UriKind.Relative),
                    "application/xml",
                    "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility");
            }
        }

        private static IEnumerable<XElement> GetShapesXml(XDocument pageXML)
        {
            return GetXElementsByName(pageXML, "Shape");
        }

        private IEnumerable<string> GetPrefixedTextsFromShapes(IEnumerable<XElement> shapesXML, string searchText)
        {
            List<string> result = new();

            foreach (var shape in shapesXML)
            {
                var texts = GetXElementsByName(shape, "Text");

                if (texts.Any())
                {
                    var textValues = texts.Select(t => t.Value.TrimEnd());
                    var filteredTexts = textValues.Where(t => t.StartsWith(searchText));
                    result.AddRange(filteredTexts);
                }
            }

            return result;
        }

        private IEnumerable<PackagePart> GetPages()
        {
            PackagePart documentPart = GetDocument();
            PackagePart pagesPart = GetPagesCollection(documentPart);
            _pagesNames = GetPagesNames(pagesPart);
            IEnumerable<PackagePart> pages = GetPages(pagesPart);
            return pages;
        }

        private IEnumerable<PackagePart> GetPages(PackagePart pagesPart)
        {
            return GetPackageParts(pagesPart, Relationships.Page);
        }

        private PackagePart GetPagesCollection(PackagePart documentPart)
        {
            return GetPackagePart(documentPart, Relationships.Pages);
        }

        private PackagePart GetDocument()
        {
            return GetPackagePart(Relationships.Document);
        }

        private IEnumerable<string> GetPagesNames(PackagePart pagesPart)
        {
            XDocument pagesXml = GetXMLFromPart(pagesPart);
            IEnumerable<XElement> pageElements = GetXElementsByName(pagesXml, "Page");
            IEnumerable<string> names = pageElements.Select(p => p.Attribute("Name")?.Value.Trim());
            return names;
        }

        private static void IteratePackageParts(Package filePackage)
        {

            // Get all of the package parts contained in the package
            // and then write the URI and content type of each one to the console.
            PackagePartCollection packageParts = filePackage.GetParts();
            foreach (PackagePart part in packageParts)
            {
                Debug.WriteLine("Package part URI: {0}", part.Uri);
                Debug.WriteLine("Content type: {0}", part.ContentType.ToString());
            }
        }

        private static Package OpenPackage(string filePath)
        {
            FileInfo visioFile = new(filePath);

            if (!visioFile.Exists)
            {
                throw new FileNotFoundException(filePath);
            }

            var visioPackage = Package.Open(
                filePath,
                FileMode.Open,
                FileAccess.ReadWrite);

            return visioPackage;
        }

        private PackagePart GetPackagePart(string relationship)
        {
            PackageRelationship packageRel = _visioPackage.GetRelationshipsByType(relationship).FirstOrDefault();
            PackagePart part = null;

            if (packageRel != null)
            {
                Uri docUri = PackUriHelper.ResolvePartUri(new Uri("/", UriKind.Relative), packageRel.TargetUri);
                part = _visioPackage.GetPart(docUri);
            }

            return part;
        }

        private PackagePart GetPackagePart(PackagePart sourcePart, string relationship)
        {
            // This gets only the first PackagePart that shares the relationship
            // with the PackagePart passed in as an argument. You can modify the code
            // here to return a different PackageRelationship from the collection.
            PackageRelationship packageRel = sourcePart.GetRelationshipsByType(relationship).FirstOrDefault();
            PackagePart relatedPart = null;

            if (packageRel != null)
            {
                // Use the PackUriHelper class to determine the URI of PackagePart
                // that has the specified relationship to the PackagePart passed in
                // as an argument.
                Uri partUri = PackUriHelper.ResolvePartUri(sourcePart.Uri, packageRel.TargetUri);
                relatedPart = _visioPackage.GetPart(partUri);
            }

            return relatedPart;
        }

        private IEnumerable<PackagePart> GetPackageParts(PackagePart sourcePart, string relationship)
        {
            var packageRelCollection = sourcePart.GetRelationshipsByType(relationship).OrderBy(i => i.Id);

            List<PackagePart> result = new(packageRelCollection.Count());

            foreach (var packageRel in packageRelCollection)
            {
                Uri partUri = PackUriHelper.ResolvePartUri(sourcePart.Uri, packageRel.TargetUri);
                var relatedPart = _visioPackage.GetPart(partUri);
                result.Add(relatedPart);
            }

            return result;
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

        private static IEnumerable<XElement> GetXElementsByName(XDocument packagePart, string elementType)
        {
            // Construct a LINQ query that selects elements by their element type.
            IEnumerable<XElement> elements =
                from element in packagePart.Descendants()
                where element.Name.LocalName == elementType
                select element;
            // Return the selected elements to the calling code.
            return elements.DefaultIfEmpty(null);
        }

        private static IEnumerable<XElement> GetXElementsByName(XElement packagePart, string elementType)
        {
            // Construct a LINQ query that selects elements by their element type.
            IEnumerable<XElement> elements =
                from element in packagePart.Descendants()
                where element.Name.LocalName == elementType
                select element;

            // Return the selected elements to the calling code.
            return elements ?? Array.Empty<XElement>();
        }

        private static XElement GetXElementByAttribute(IEnumerable<XElement> elements, string attributeName, string attributeValue)
        {
            // Construct a LINQ query that selects elements from a group
            // of elements by the value of a specific attribute.
            IEnumerable<XElement> selectedElements =
                from el in elements
                where el.Attribute(attributeName)?.Value == attributeValue
                select el;
            // If there aren't any elements of the specified type
            // with the specified attribute value in the document,
            // return null to the calling code.
            return selectedElements.DefaultIfEmpty(null).FirstOrDefault();
        }

        private static void SaveXDocumentToPart(PackagePart packagePart, XDocument partXML)
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

        private void RecalcDocument()
        {
            // Get the Custom File Properties part from the package and
            // and then extract the XML from it.
            PackagePart customPart = GetPackagePart("http://schemas.openxmlformats.org/officeDocument/2006/relationships/" +
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

        private static XDocument CreateCustomUI()
        {
            // Add a new Custom User Interface document part to the package.
            // This code adds a new CUSTOM tab to the ribbon for this
            // document. The tab has one group that contains one button.
            XNamespace customUINS =
                "http://schemas.microsoft.com/office/2006/01/customui";
            XDocument customUIXDoc = new XDocument(
                new XDeclaration("1.0", "utf-8", "true"),
                new XElement(customUINS + "customUI",
                    new XElement(customUINS + "ribbon",
                        new XElement(customUINS + "tabs",
                            new XElement(customUINS + "tab",
                                new XAttribute("id", "customTab"),
                                new XAttribute("label", "CUSTOM"),
                                new XElement(customUINS + "group",
                                    new XAttribute("id", "customGroup"),
                                    new XAttribute("label", "Custom Group"),
                                    new XElement(customUINS + "button",
                                        new XAttribute("id", "customButton"),
                                        new XAttribute("label", "Custom Button"),
                                        new XAttribute("size", "large"),
                                        new XAttribute("imageMso", "HappyFace")
                                    )
                                )
                            )
                        )
                    )
                )
            );
            return customUIXDoc;
        }

        private static void CreateNewPackagePart(Package filePackage, XDocument partXML, Uri packageLocation, string contentType, string relationship)
        {
            // Need to check first to see whether the part exists already.
            if (!filePackage.PartExists(packageLocation))
            {
                // Create a new blank package part at the specified URI 
                // of the specified content type.
                PackagePart newPackagePart = filePackage.CreatePart(packageLocation,
                    contentType);
                // Create a stream from the package part and save the 
                // XML document to the package part.
                using (Stream partStream = newPackagePart.GetStream(FileMode.Create,
                    FileAccess.ReadWrite))
                {
                    partXML.Save(partStream);
                }
            }
            // Add a relationship from the file package to this
            // package part. You can also create relationships
            // between an existing package part and a new part.
            filePackage.CreateRelationship(packageLocation,
                TargetMode.Internal,
                relationship);
        }

    }
}