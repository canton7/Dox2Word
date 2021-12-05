using System;
using System.Collections.Generic;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Serialization;
using Dox2Word.Logging;
using Dox2Word.Model;

namespace Dox2Word.Parser.Models
{
    public abstract class DocTitleCmdGroup : IXmlSerializable
    {
        private static readonly Logger logger = Logger.Instance;

        public List<object> Parts { get; } = new();

        protected virtual object? ParseElement(XmlReader reader)
        {
            if (DocEmptyParser.TryLookup(reader.Name, out string? result))
            {
                return result;
            }

            return reader.Name switch
            {
                "ulink" => Load<DocUrlLink>(reader),
                "bold" => Load<DocMarkup>(reader, x => x.Properties = new(TextRunFormat.Bold)),
                "s" or "strike" => Load<DocMarkup>(reader, x => x.Properties = new(TextRunFormat.Strikethrough)),
                "underline" => Load<DocMarkup>(reader, x => x.Properties = new(TextRunFormat.Underline)),
                "emphasis" => Load<DocMarkup>(reader, x => x.Properties = new(TextRunFormat.Italic)),
                "computeroutput" => Load<DocMarkup>(reader, x => x.Properties = new(TextRunFormat.Monospace)),
                "subscript" => Load<DocMarkup>(reader, x => x.Properties = new(TextRunVerticalPosition.Subscript)),
                "superscript" => Load<DocMarkup>(reader, x => x.Properties = new(TextRunVerticalPosition.Superscript)),
                "center" => Load<DocMarkup>(reader, x => x.Center = true),
                "small" => Load<DocMarkup>(reader, x => x.Properties = TextRunProperties.Small),
                "del" => Load<DocMarkup>(reader, x => x.Properties = new(TextRunFormat.Strikethrough)),
                "ins" => Load<DocMarkup>(reader, x => x.Properties = new(TextRunFormat.Underline)),
                "htmlonly" or "manonly" => null,
                "xmlonly" => reader.ReadElementContentAsString().Trim(' ', '\n'), // We don't actually get these...
                "rtfonly" or "latexonly" or "docbookonly" => null,
                "image" => Load<Image>(reader),
                "dot" => Load<Dot>(reader),
                "msc" => Unsupported("msc"),
                "plantuml" => Unsupported("plantuml"), // TODO: plantuml
                "anchor" => Load<DocAnchor>(reader),
                "formula" => reader.ReadElementContentAsString().Trim(' ', '\n'), // TODO: HTML doesn't handle these well, but maybe we can use plantuml here?
                "ref" => Load<Ref>(reader),
                "emoji" => Load<DocEmoji>(reader),
                "linebreak" => "\n",
                _ => LoadXmlElement(reader),
            };
        }

        public void ReadXml(XmlReader reader)
        {
            int startingDepth = reader.Depth;
            this.ReadAttributes(reader);

            // We only care about whitespace when it sits between two child elements
            string? whitespace = null;

            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    // Whitespace is significant if it comes before a child (e.g. <bold>foo</bold> <bold>bar</bold>)
                    if (!string.IsNullOrEmpty(whitespace))
                    {
                        this.ProcessPart(whitespace!);
                    }
                    reader.MoveToContent();
                    var subReader = reader.ReadSubtree();
                    subReader.Read();
                    object? result = this.ParseElement(subReader);
                    if (result != null)
                    {
                        this.ProcessPart(result);
                    }
                    else
                    {
                        subReader.Close();
                    }
                }

                whitespace = null;

                switch (reader.NodeType)
                {
                    case XmlNodeType.Text:
                    case XmlNodeType.CDATA:
                        this.ProcessPart(reader.Value.TrimStart(' ', '\n').TrimEnd('\n'));
                        break;

                    case XmlNodeType.SignificantWhitespace: // Whitespace within xml:space="preserve"
                        this.ProcessPart(reader.Value);
                        break;

                    case XmlNodeType.Whitespace:
                        whitespace = reader.Value.Substring(reader.Value.LastIndexOf('\n') + 1);
                        break;

                    case XmlNodeType.EntityReference:
                        logger.Warning($"Unknown entity reference {reader.Name}");
                        break;

                    case XmlNodeType.EndElement when reader.Depth == startingDepth:
                        reader.ReadEndElement();
                        return;

                    default:
                        break;
                }
            }
        }

        private void ProcessPart(object part)
        {
            this.Parts.Add(part);
        }

        protected virtual void ReadAttributes(XmlReader reader) { }

        protected static T Load<T>(XmlReader reader, Action<T>? configurer = null)
        {
            var instance = (T)SerializerCache.Get<T>(reader.Name).Deserialize(reader);
            configurer?.Invoke(instance);
            return instance;
        }

        protected static XmlElement LoadXmlElement(XmlReader reader)
        {
            var doc = new XmlDocument();
            doc.Load(reader);
            return doc.DocumentElement;
        }

        protected static object? Unsupported(string name)
        {
            logger.Unsupported($"Doxygen command {name} is not supported. Ignoring");
            return null;
        }

        public static object? Do(Action a)
        {
            a();
            return null;
        }

        public void WriteXml(XmlWriter writer) => throw new NotSupportedException();
        public XmlSchema? GetSchema() => null;

    }
}
