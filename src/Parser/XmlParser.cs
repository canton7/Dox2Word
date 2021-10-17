using System;
using System.Collections.Generic;
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
                        ReturnDescription = ParaToParagraph(member.DetailedDescription?.Para.SelectMany(x => x.Parts)
                            .OfType<DocSimpleSect>()
                            .FirstOrDefault(x => x.Kind == DoxSimpleSectKind.Return)?.Para),
                        ArgsString = member.ArgsString ?? "",
                    };
                    function.FunctionParameters.AddRange(ParseFunctionParameters(member));
                    group.Functions.Add(function);
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
            }

            return group;
        }

        private static IEnumerable<FunctionParameter> ParseFunctionParameters(MemberDef member)
        {
            foreach (var param in member.Params)
            {
                // Find its docs...
                var descriptionPara = member.DetailedDescription?.Para.SelectMany(x => x.ParameterLists)
                    .Where(x => x.Kind == DoxParamListKind.Param)
                    .SelectMany(x => x.ParameterItems)
                    .FirstOrDefault(x => x.ParameterNameList.Select(x => x.ParameterName).Contains(param.DeclName))
                    ?.ParameterDescription.Para.FirstOrDefault();

                var functionParameter = new FunctionParameter()
                {
                    Name = param.DeclName ?? "",
                    Type = LinkedTextToString(param.Type) ?? "",
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
                var variable = new ClassVariable()
                { 
                    Name = member.Name ?? "",
                    Type = LinkedTextToString(member.Type) ?? "",
                    Definition = member.Definition ?? "",
                    Descriptions = ParseDescriptions(member),
                };
                cls.Variables.Add(variable);
            }

            return cls;
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
                BriefDescription = ParaToParagraph(member.BriefDescription?.Para.FirstOrDefault()),
            };
            descriptions.DetailedDescrpition.AddRange(ParasToParagraphs(member.DetailedDescription?.Para));
            return descriptions;
        }

        private static IEnumerable<Paragraph> ParasToParagraphs(IEnumerable<DocPara>? paras)
        {
            if (paras == null)
                return Enumerable.Empty<Paragraph>();

            return paras.Select(x => ParaToParagraph(x)).Where(x => x.Count > 0);
        }

        private static Paragraph ParaToParagraph(DocPara? para)
        {
            var paragraph = new Paragraph();

            if (para == null)
                return paragraph;

            foreach (var run in Parse(para))
            {
                paragraph.Add(run);
            }

            IEnumerable<ITextRun> Parse(DocPara para)
            {
                foreach (object? part in para.Parts)
                {
                    ITextRun? textRun = part switch
                    {
                        string s => new LiteralTextRun(s),
                        BoldMarkup b => new TextRun(TextRunFormat.Bold, Parse(b)),
                        ItalicMarkup i => new TextRun(TextRunFormat.Italic, Parse(i)),
                        MonospaceMarkup m => new TextRun(TextRunFormat.Monospace, Parse(m)),
                        XmlElement e => new LiteralTextRun(e.InnerText),
                        // Ignore other types
                        _ => null,
                    };
                    if (textRun != null)
                    {
                        yield return textRun;
                    }
                }
            }

            return paragraph;
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
