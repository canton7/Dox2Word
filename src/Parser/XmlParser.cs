using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;
using Dox2Word.Model;
using Dox2Word.Parser.Models;

namespace Dox2Word.Parser
{
    public class XmlParser
    {
        private readonly string basePath;

        public XmlParser(string basePath)
        {
            this.basePath = basePath;
        }

        public Project Parse()
        {
            string indexFile = Path.Join(this.basePath, "index.xml");
            var index = Parse<DoxygenIndex>(indexFile);

            var project = new Project();

            foreach (var group in index.Compounds.Where(x => x.Kind == CompoundKind.Group).OrderBy(x => x.Name))
            {
                project.Groups.Add(this.ParseGroup(group.RefId));
            }

            return project;
        }

        private Group ParseGroup(string refId)
        {
            var compoundDef = this.ParseDoxygenFile(refId);

            var group = new Group()
            {
                Name = compoundDef.Title,
                Descriptions = ParseDescriptions(compoundDef),
            };
            group.SubGroups.AddRange(compoundDef.InnerGroups.Select(x => this.ParseGroup(x.RefId)));
            group.Files.AddRange(compoundDef.InnerFiles.Select(x => x.Name));
            group.Classes.AddRange(compoundDef.InnerClasses.Select(x => this.ParseClass(x.RefId)));

            var members = compoundDef.Sections.SelectMany(x => x.Members);
            foreach (var member in members)
            {
                if (member.Kind == DoxMemberKind.Function)
                {
                    var function = new Function()
                    {
                        Name = member.Name,
                        Descriptions = ParseDescriptions(member),
                        ReturnType = LinkedTextToString(member.Type) ?? "",
                        ReturnDescription = ParseReturnDescription(member),
                        ArgsString = member.ArgsString ?? "",
                    };
                    function.Parameters.AddRange(ParseParameters(member));
                    group.Functions.Add(function);
                }
                else if (member.Kind == DoxMemberKind.Define)
                {
                    var macro = new Macro()
                    { 
                        Name = member.Name,
                        Descriptions = ParseDescriptions(member),
                        ReturnDescription = ParseReturnDescription(member),
                        Initializer = LinkedTextToString(member.Initializer) ?? "",
                    };
                    macro.Parameters.AddRange(ParseParameters(member));
                    group.Macros.Add(macro);
                }
                else if (member.Kind == DoxMemberKind.Typedef)
                {
                    var typedef = new Typedef()
                    {
                        Name = member.Name,
                        Type = LinkedTextToString(member.Type) ?? "",
                        Definition = member.Definition ?? "",
                        Descriptions = ParseDescriptions(member),
                    };
                    group.Typedefs.Add(typedef);
                }
                else if (member.Kind == DoxMemberKind.Variable)
                {
                    group.GlobalVariables.Add(ParseVariable(member));
                }
            }

            return group;
        }

        private static IEnumerable<Parameter> ParseParameters(MemberDef member)
        {
            foreach (var param in member.Params)
            {
                string name = param.DeclName ?? param.DefName ?? "";

                // Find its docs...
                var descriptionPara = member.DetailedDescription?.Para.SelectMany(x => x.ParameterLists)
                    .Where(x => x.Kind == DoxParamListKind.Param)
                    .SelectMany(x => x.ParameterItems)
                    .FirstOrDefault(x => x.ParameterNameList.Select(x => x.ParameterName).Contains(name))
                    ?.ParameterDescription.Para.FirstOrDefault();

                var functionParameter = new Parameter()
                {
                    Name = name,
                    Type = LinkedTextToString(param.Type),
                    Description = ParaToParagraph(descriptionPara),
                };

                yield return functionParameter;
            }
        }

        private Class ParseClass(string refId)
        {
            var compoundDef = this.ParseDoxygenFile(refId);

            if (compoundDef.Kind != CompoundKind.Struct)
                throw new ParserException($"Don't konw how to parse class kind {compoundDef.Kind} in {refId}");

            var cls = new Class()
            {
                Name = compoundDef.CompoundName ?? "",
                Descriptions = ParseDescriptions(compoundDef),
            };

            var members = compoundDef.Sections.SelectMany(x => x.Members)
                .Where(x => x.Kind == DoxMemberKind.Variable);
            foreach (var member in members)
            {
                cls.Variables.Add(ParseVariable(member));
            }

            return cls;
        }

        private static Variable ParseVariable(MemberDef member)
        {
            var variable = new Variable()
            {
                Name = member.Name ?? "",
                Type = LinkedTextToString(member.Type) ?? "",
                Definition = member.Definition ?? "",
                Descriptions = ParseDescriptions(member),
            };
            return variable;
        }

        private static Paragraph ParseReturnDescription(MemberDef member)
        {
            return ParaToParagraph(member.DetailedDescription?.Para.SelectMany(x => x.Parts)
                .OfType<DocSimpleSect>()
                .FirstOrDefault(x => x.Kind == DoxSimpleSectKind.Return)?.Para);
        }

        private static string? LinkedTextToString(LinkedText? linkedText)
        {
            if (linkedText == null)
                return null;

            return string.Join(" ", linkedText.Type.Select(x =>
                x switch
                {
                    string s => s,
                    RefText r => r.Name,
                    _ => throw new ParserException($"Unknown element in LinkedText: {x}"),
                }));
        }

        private static Descriptions ParseDescriptions(IDoxDescribable member)
        {
            var descriptions = new Descriptions()
            {
                BriefDescription = ParasToParagraph(member.BriefDescription?.Para),
            };
            descriptions.DetailedDescription.AddRange(ParasToParagraphs(member.DetailedDescription?.Para));
            return descriptions;
        }

        private static Paragraph ParaToParagraph(DocPara? para)
        {
            return ParaToParagraphs(para).FirstOrDefault() ?? new Paragraph();
        }

        private static Paragraph ParasToParagraph(IEnumerable<DocPara>? paras)
        {
            return ParasToParagraphs(paras).FirstOrDefault() ?? new Paragraph();
        }

        private static IEnumerable<Paragraph> ParasToParagraphs(IEnumerable<DocPara>? paras)
        {
            if (paras == null)
                return Enumerable.Empty<Paragraph>();

            return paras.SelectMany(x => ParaToParagraphs(x)).Where(x => x.Count > 0);
        }

        private static List<Paragraph> ParaToParagraphs(DocPara? para)
        {
            var paragraphs = new List<Paragraph>();

            if (para == null)
                return paragraphs;

            paragraphs.Add(new Paragraph());
            Parse(paragraphs, para, TextRunFormat.None);

            static void Parse(List<Paragraph> paragraphs, DocPara? para, TextRunFormat format)
            {
                void NewParagraph(ParagraphType type = ParagraphType.Normal) => paragraphs.Add(new Paragraph(type));

                void Add(ITextRun textRun) => paragraphs[paragraphs.Count - 1].Add(textRun);
                void AddTextRun(string text, TextRunFormat format) => Add(new TextRun(text, format));

                if (para == null)
                    return;

                foreach (object? part in para.Parts)
                {
                    switch (part)
                    {
                        case string s:
                            AddTextRun(s, format);
                            break;
                        case DocSimpleSect s when s.Kind == DoxSimpleSectKind.Warning:
                            NewParagraph(ParagraphType.Warning);
                            Parse(paragraphs, s.Para, format);
                            NewParagraph();
                            break;
                        case OrderedList o:
                            ParseList(o, ListTextRunType.Number, format);
                            break;
                        case UnorderedList u:
                            ParseList(u, ListTextRunType.Bullet, format);
                            break;
                        case BoldMarkup b:
                            Parse(paragraphs, b, format | TextRunFormat.Bold);
                            break;
                        case ItalicMarkup i:
                            Parse(paragraphs, i, format | TextRunFormat.Italic);
                            break;
                        case MonospaceMarkup m:
                            Parse(paragraphs, m, format | TextRunFormat.Monospace);
                            break;
                        case XmlElement e:
                            AddTextRun(e.InnerText, format);
                            break;
                    };
                }

                void ParseList(DocList docList, ListTextRunType type, TextRunFormat format)
                {
                    var list = new ListTextRun(type);
                    Add(list);
                    foreach (var item in docList.Items)
                    {
                        var runItem = new ListTextRunItem();
                        // It could be that the para contains a warning or something which will add
                        // another paragraph. In this case, we'll just ignore it.
                        var paragraphList = new List<Paragraph>() { runItem };
                        foreach (var para in item.Paras)
                        {
                            Parse(paragraphList, para, format);
                        }
                        list.Items.Add(runItem);
                    }
                }
            }

            return paragraphs;
        }

        private CompoundDef ParseDoxygenFile(string refId)
        {
            string filePath = Path.Combine(this.basePath, refId + ".xml");
            var file = Parse<DoxygenFile>(filePath);
            if (file.CompoundDefs.Count != 1)
                throw new ParserException($"File {filePath}: expected 1 compoundDef, got {file.CompoundDefs.Count}");
            return file.CompoundDefs[0];
        }

        private static class SerializerCache<T>
        {
            public static readonly XmlSerializer Instance = new XmlSerializer(typeof(T));
        }
        private static T Parse<T>(string filePath)
        {
            using (var stream = File.OpenRead(filePath))
            {
                return (T)SerializerCache<T>.Instance.Deserialize(stream)!;
            }
        }
    }
}
