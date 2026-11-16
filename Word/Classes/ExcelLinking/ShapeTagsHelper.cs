// ShapeTagsHelper.cs (.NET Framework 4.8 / C# 7.3)
// Word shapes don't have a native Tags collection like PowerPoint shapes.
// This helper uses CustomXMLParts to store shape-specific tags in the document.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;

namespace OfficePragmatics_Word.ExcelLinking
{
    /// <summary>
    /// Provides tag storage for Word shapes using CustomXMLParts.
    /// Word shapes don't have a native Tags collection, so we use a document-level
    /// XML store keyed by shape ID.
    /// </summary>
    internal static class ShapeTagsHelper
    {
        private const string CustomXmlNamespace = "http://officepragmatics.com/word/shapetags";
        private const string RootElementName = "ShapeTags";
        private const string ShapeElementName = "Shape";
        private const string TagElementName = "Tag";
        private const string IdAttribute = "id";
        private const string NameAttribute = "name";

        // Cache for loaded tags - keyed by document to support multiple docs
        private static readonly Dictionary<int, Dictionary<int, Dictionary<string, string>>> _documentCache =
            new Dictionary<int, Dictionary<int, Dictionary<string, string>>>();

        /// <summary>
        /// Gets a tag value from a shape.
        /// </summary>
        public static string GetTag(Word.Shape shape, string tagName)
        {
            if (shape == null || string.IsNullOrEmpty(tagName))
                return null;

            try
            {
                int shapeId = shape.ID;
                Word.Document doc = shape.Parent as Word.Document;
                if (doc == null)
                {
                    // Shape might be in a header/footer - get document from range
                    try
                    {
                        doc = shape.Anchor?.Document;
                    }
                    catch { }
                }
                if (doc == null) return null;

                var shapeTags = GetOrLoadShapeTags(doc, shapeId);
                if (shapeTags != null && shapeTags.TryGetValue(tagName, out string value))
                    return value;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting tag '{tagName}': {ex.Message}");
            }

            return null;
        }

        /// <summary>
        /// Sets a tag value on a shape.
        /// </summary>
        public static void SetTag(Word.Shape shape, string tagName, string value)
        {
            if (shape == null || string.IsNullOrEmpty(tagName))
                return;

            try
            {
                int shapeId = shape.ID;
                Word.Document doc = shape.Parent as Word.Document;
                if (doc == null)
                {
                    try
                    {
                        doc = shape.Anchor?.Document;
                    }
                    catch { }
                }
                if (doc == null) return;

                var shapeTags = GetOrLoadShapeTags(doc, shapeId);
                if (shapeTags == null)
                {
                    shapeTags = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    EnsureDocumentCache(doc)[shapeId] = shapeTags;
                }

                if (string.IsNullOrEmpty(value))
                    shapeTags.Remove(tagName);
                else
                    shapeTags[tagName] = value;

                SaveTagsToDocument(doc);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error setting tag '{tagName}': {ex.Message}");
            }
        }

        #region Private Methods

        private static Dictionary<int, Dictionary<string, string>> EnsureDocumentCache(Word.Document doc)
        {
            int docHash = doc.GetHashCode();
            if (!_documentCache.TryGetValue(docHash, out var docCache))
            {
                docCache = new Dictionary<int, Dictionary<string, string>>();
                _documentCache[docHash] = docCache;
            }
            return docCache;
        }

        private static Dictionary<string, string> GetOrLoadShapeTags(Word.Document doc, int shapeId)
        {
            var docCache = EnsureDocumentCache(doc);

            // If cache is empty, load from document
            if (docCache.Count == 0)
            {
                LoadTagsFromDocument(doc);
            }

            docCache.TryGetValue(shapeId, out var shapeTags);
            return shapeTags;
        }

        private static void LoadTagsFromDocument(Word.Document doc)
        {
            if (doc == null) return;

            var docCache = EnsureDocumentCache(doc);
            docCache.Clear();

            try
            {
                Office.CustomXMLPart customPart = FindCustomXmlPart(doc);
                if (customPart == null) return;

                string xml = customPart.XML;
                if (string.IsNullOrEmpty(xml)) return;

                XDocument xdoc = XDocument.Parse(xml);
                XNamespace ns = CustomXmlNamespace;

                foreach (var shapeElement in xdoc.Descendants(ns + ShapeElementName))
                {
                    string idStr = shapeElement.Attribute(IdAttribute)?.Value;
                    if (!int.TryParse(idStr, out int shapeId)) continue;

                    var shapeTags = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                    foreach (var tagElement in shapeElement.Elements(ns + TagElementName))
                    {
                        string name = tagElement.Attribute(NameAttribute)?.Value;
                        string value = tagElement.Value;
                        if (!string.IsNullOrEmpty(name))
                            shapeTags[name] = value ?? string.Empty;
                    }

                    if (shapeTags.Count > 0)
                        docCache[shapeId] = shapeTags;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading tags from document: {ex.Message}");
            }
        }

        private static void SaveTagsToDocument(Word.Document doc)
        {
            if (doc == null) return;

            try
            {
                var docCache = EnsureDocumentCache(doc);

                // Build XML
                XNamespace ns = CustomXmlNamespace;
                var xdoc = new XDocument(
                    new XDeclaration("1.0", "utf-8", "yes"),
                    new XElement(ns + RootElementName)
                );

                foreach (var shapeKvp in docCache)
                {
                    if (shapeKvp.Value.Count == 0) continue;

                    var shapeElement = new XElement(ns + ShapeElementName,
                        new XAttribute(IdAttribute, shapeKvp.Key));

                    foreach (var tagKvp in shapeKvp.Value)
                    {
                        shapeElement.Add(new XElement(ns + TagElementName,
                            new XAttribute(NameAttribute, tagKvp.Key),
                            tagKvp.Value ?? string.Empty));
                    }

                    xdoc.Root.Add(shapeElement);
                }

                string newXml = xdoc.ToString(SaveOptions.DisableFormatting);

                // Find or create custom XML part
                Office.CustomXMLPart existingPart = FindCustomXmlPart(doc);
                if (existingPart != null)
                {
                    // Delete and recreate (can't modify in place reliably)
                    existingPart.Delete();
                }

                // Add new part
                doc.CustomXMLParts.Add(newXml);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error saving tags to document: {ex.Message}");
            }
        }

        private static Office.CustomXMLPart FindCustomXmlPart(Word.Document doc)
        {
            if (doc == null) return null;

            try
            {
                foreach (Office.CustomXMLPart part in doc.CustomXMLParts)
                {
                    try
                    {
                        if (part.NamespaceURI == CustomXmlNamespace)
                            return part;
                    }
                    catch { }
                }
            }
            catch { }

            return null;
        }

        #endregion
    }
}