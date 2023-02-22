using System.IO.Packaging;
using System.Xml.Linq;

namespace Helpers
{
    public class VisioHelper
    {
        private IEnumerable<PackagePart>? _pages;
        private IEnumerable<string>? _pagesNames;
        private Package? _visioPackage;

        public Dictionary<string, IEnumerable<string>> SearchForPrefixTextInVisioFile(string fileName, string prefixText)
        {
            Dictionary<string, IEnumerable<string>> result = new();

            using (_visioPackage = OpenPackage(fileName))
            {
                GetPages();

                int currentPageIndex = 0;

                foreach (PackagePart page in _pages!)
                {
                    IEnumerable<string> texts = GetPagePrefixedTexts(prefixText, page);

                    if (texts.Any())
                    {
                        string pageName = _pagesNames!.ElementAt(currentPageIndex++);
                        result.Add(pageName, texts);
                    }
                }

                return result;
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

            if (visioPackage == null)
                throw new NullReferenceException($"Could not open a package from file {filePath}.");

            return visioPackage;
        }

        private PackagePart GetDocumentPart()
        {
            return GetPackagePart(Relationships.Document);
        }

        private PackagePart GetPackagePart(string relationship)
        {
            PackageRelationship? packageRel = _visioPackage!.GetRelationshipsByType(relationship).FirstOrDefault();

            PackagePart? part = null;

            if (packageRel != null)
            {
                Uri sourceUri = new("/", UriKind.Relative);
                Uri targetUri = packageRel.TargetUri;
                part = GetPackagePart(sourceUri, targetUri);
            }

            return part ?? throw new ArgumentException("No relationship was found.");
        }

        private PackagePart GetPackagePart(PackagePart sourcePart, string relationship)
        {
            PackageRelationship? packageRel = sourcePart.GetRelationshipsByType(relationship).FirstOrDefault();

            PackagePart? part = null;

            if (packageRel != null)
            {
                Uri sourceUri = sourcePart.Uri;
                Uri targetUri = packageRel.TargetUri;
                part = GetPackagePart(sourceUri, targetUri);
            }

            return part ?? throw new ArgumentException("No package part was found.");
        }

        private PackagePart GetPackagePart(Uri sourceUri, Uri targetUri)
        {
            Uri? partUri = PackUriHelper.ResolvePartUri(sourceUri, targetUri);
            var relatedPart = _visioPackage!.GetPart(partUri);
            return relatedPart;
        }

        private IEnumerable<PackagePart> GetPackageParts(PackagePart sourcePart, string relationship)
        {
            var packageRelCollection = sourcePart.GetRelationshipsByType(relationship).OrderBy(i => i.Id);

            List<PackagePart> result = new(packageRelCollection.Count());

            foreach (var packageRel in packageRelCollection)
            {
                Uri partUri = PackUriHelper.ResolvePartUri(sourcePart.Uri, packageRel.TargetUri);
                var relatedPart = _visioPackage!.GetPart(partUri);
                result.Add(relatedPart);
            }

            return result;
        }

        private IEnumerable<string> GetPagePrefixedTexts(string prefixText, PackagePart page)
        {
            XDocument pageXML = GetXMLFromPart(page);
            IEnumerable<XElement> shapesXML = GetShapesXml(pageXML);
            IEnumerable<string> texts = GetPrefixedTextsFromShapes(shapesXML, prefixText);
            return texts;
        }

        private static IEnumerable<XElement> GetShapesXml(XDocument pageXML)
        {
            return GetXElementsByName(pageXML, "Shape");
        }

        private static IEnumerable<XElement> GetXElementsByName(XContainer packagePart, string elementType)
        {
            IEnumerable<XElement> elements =
                from element in packagePart.Descendants()
                where element.Name.LocalName == elementType
                select element;

            return elements ?? Array.Empty<XElement>();
        }

        private static XDocument GetXMLFromPart(PackagePart packagePart)
        {
            Stream partStream = packagePart.GetStream();
            XDocument partXml = XDocument.Load(partStream);
            return partXml ?? throw new ArgumentException("Could not get xml part from package part."); ;
        }

        private void GetPages()
        {
            PackagePart documentPart = GetDocumentPart();
            PackagePart pagesCollection = GetPagesCollection(documentPart);
            _pagesNames = GetPagesNames(pagesCollection);
            IEnumerable<PackagePart> pages = GetPagesFromPagesCollection(pagesCollection);
            _pages = pages;
            ThrowOnNoPages();
        }

        private PackagePart GetPagesCollection(PackagePart documentPart)
        {
            return GetPackagePart(documentPart, Relationships.Pages);
        }

        private IEnumerable<PackagePart> GetPagesFromPagesCollection(PackagePart pagesPart)
        {
            return GetPackageParts(pagesPart, Relationships.Page);
        }

        private IEnumerable<string> GetPagesNames(PackagePart pagesPart)
        {
            XDocument pagesXml = GetXMLFromPart(pagesPart);
            IEnumerable<XElement> pageElements = GetXElementsByName(pagesXml, "Page");
            IEnumerable<XAttribute?> elementsNames = pageElements.Where(p => p.Attribute("Background") == null).Select(p => p.Attribute("Name")).Where(e => e != null);
            IEnumerable<string> names = elementsNames.Select(p => p!.Value.ToString().Trim());
            return names;
        }

        private IEnumerable<string> GetPrefixedTextsFromShapes(IEnumerable<XElement> shapesXML, string searchText)
        {
            List<string> result = new();

            foreach (var shape in shapesXML)
            {
                var texts = GetXElementsByName(shape, "Text");

                if (texts.Any())
                {
                    var textValues = texts.Select(t => t.Value.Trim());
                    var filteredTexts = textValues.Where(t => t.StartsWith(searchText));
                    result.AddRange(filteredTexts);
                }
            }

            return result;
        }

        private void ThrowOnNoPages()
        {
            if (_pages == null)
            {
                throw new ArgumentException("No pages were found.");
            }

            if (_pagesNames == null)
            {
                throw new ArgumentException("No pages names were found.");
            }
        }
    }
}