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
            string filePath = Path.Combine(this.basePath, refId + ".xml");
            var file = Parse<DoxygenFile>(filePath);

            if (file.CompoundDefs.Count != 1)
                throw new ParserException($"File {filePath}: expected 1 compountDef, got {file.CompoundDefs.Count}");
            var compoundDef = file.CompoundDefs[0];

            var group = new Group()
            {
                Name = compoundDef.Title,
            };

            foreach (var innerGroup in compoundDef.InnerGroups)
            {
                group.SubGroups.Add(this.ParseGroup(innerGroup.RefId));
            }

            foreach (var section in compoundDef.Sections)
            {
                foreach (var member in section.Members)
                {
                    var function = new Function()
                    {
                        Name = member.Name,
                        BriefDescription = ParaToParagraph(member.BriefDescription?.Para.FirstOrDefault()),
                        ReturnType = LinkedTextToString(member.Type) ?? "",
                        ArgsString = member.ArgsString ?? "",
                    };
                    function.DetailedDescription.AddRange(ParasToParagraphs(member.DetailedDescription?.Para));
                    function.FunctionParameters.AddRange(ParseFunctionParameters(member));
                    group.Functions.Add(function);
                }
            }

            return group;
        }

        private static IEnumerable<FunctionParameter> ParseFunctionParameters(MemberDef member)
        {
            foreach (var param in member.Params)
            {
                // Find its docs...
                var descriptionPara = member.DetailedDescription?.Para.SelectMany(x => x.Parts)
                    .OfType<DocParamList>()
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

            foreach (object? part in para.Parts)
            {
                var textRun = part switch
                {
                    string s => new TextRun(s),
                    XmlElement e => new TextRun(e.InnerText, e.Name),
                    // Ignore other types
                    _ => null,
                };
                if (textRun != null)
                {
                    paragraph.Add(textRun);
                }
            }

            return paragraph;
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
