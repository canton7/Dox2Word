﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Serialization;
using Dox2Word.Logging;
using Dox2Word.Model;
using Dox2Word.Parser.Models;

namespace Dox2Word.Parser
{
    public class XmlParser
    {
        private static readonly Logger logger = Logger.Instance;
        private readonly string basePath;

        public XmlParser(string basePath)
        {
            this.basePath = basePath;
        }

        public Project Parse()
        {
            var project = new Project();

            try
            {
                string doxyfileFile = Path.Combine(this.basePath, "Doxyfile.xml");
                var doxyfile = Parse<DoxygenFile>(doxyfileFile);
                project.Options.AddRange(this.ParseOptions(doxyfile));

                string indexFile = Path.Combine(this.basePath, "index.xml");
                var index = Parse<DoxygenIndex>(indexFile);

                // Discover the root groups
                var groupCompoundDefs = index.Compounds.Where(x => x.Kind == CompoundKind.Group)
                    .ToDictionary(x => x.RefId, x => this.ParseDoxygenFile(x.RefId));
                var rootGroups = groupCompoundDefs.Keys.ToHashSet();
                foreach (var group in groupCompoundDefs.Values.ToList())
                {
                    foreach (var innerGroup in group.InnerGroups)
                    {
                        rootGroups.Remove(innerGroup.RefId);
                    }
                }

                project.Groups.AddRange(rootGroups.Select(x => this.ParseGroup(groupCompoundDefs, x)).OrderBy(x => x.Name));
            }
            catch (ParserException e)
            {
                logger.Error(e);
            }

            return project;
        }

        private IEnumerable<ProjectOption> ParseOptions(DoxygenFile doxyfile)
        {
            foreach (var option in doxyfile.Options)
            {
                var projectOption = new ProjectOption()
                {
                    Id = option.Id,
                };
                projectOption.TypedValue = option.Type switch
                {
                    OptionType.Bool => option.Values[0] == "YES",
                    OptionType.Int => int.Parse(option.Values[0]),
                    OptionType.String => TrimQuotes(option.Values[0]),
                    OptionType.StringList => projectOption.Values,
                };
                if (option.Type is OptionType.String or OptionType.StringList)
                {
                    projectOption.Values.AddRange(option.Values.Select(TrimQuotes));
                }
                else
                {
                    projectOption.Values.AddRange(option.Values);
                }
                yield return projectOption;
            }

            static string TrimQuotes(string value) => value.TrimStart('"').TrimEnd('"');
        }

        private Group ParseGroup(Dictionary<string, CompoundDef> groups, string refId)
        {
            var compoundDef = groups[refId];
            logger.Info($"Parsing compound {compoundDef.CompoundName}");

            var group = new Group()
            {
                Name = compoundDef.Title,
                Descriptions = ParseDescriptions(compoundDef),
            };
            group.SubGroups.AddRange(compoundDef.InnerGroups.Select(x => this.ParseGroup(groups, x.RefId)));
            group.Files.AddRange(compoundDef.InnerFiles.Select(x => x.Name));
            group.Classes.AddRange(compoundDef.InnerClasses.Select(x => this.ParseInnerClass(x.RefId)).Where(x => x != null)!);

            var members = compoundDef.Sections.SelectMany(x => x.Members);
            foreach (var member in members)
            {
                // If they didn't document it, don't include it. Doxygen itself will warn if something should have been documented
                // but wasn't.
                if (member.BriefDescription?.Para.Count is 0 or null)
                    continue;

                logger.Info($"Parsing member {member.Name}");
                switch (member.Kind)
                {
                    case DoxMemberKind.Function:
                    {
                        var function = new FunctionDoc()
                        {
                            Name = member.Name,
                            Descriptions = ParseDescriptions(member),
                            ReturnType = LinkedTextToString(member.Type) ?? "",
                            ReturnDescriptions = ParseReturnDescriptions(member),
                            Definition = member.Definition ?? "",
                            ArgsString = member.ArgsString ?? "",
                        };
                        function.Parameters.AddRange(ParseParameters(member));
                        group.Functions.Add(function);
                    }
                    break;
                    case DoxMemberKind.Define:
                    {
                        var macro = new MacroDoc()
                        {
                            Name = member.Name,
                            Descriptions = ParseDescriptions(member),
                            ReturnDescriptions = ParseReturnDescriptions(member),
                            Initializer = LinkedTextToString(member.Initializer) ?? "",
                            HasParameters = member.Params.Count > 0,
                        };
                        macro.Parameters.AddRange(ParseParameters(member));
                        group.Macros.Add(macro);
                    }
                    break;
                    case DoxMemberKind.Typedef:
                    {
                        var typedef = new TypedefDoc()
                        {
                            Name = member.Name,
                            Type = LinkedTextToString(member.Type) ?? "",
                            Definition = member.Definition ?? "",
                            Descriptions = ParseDescriptions(member),
                        };
                        group.Typedefs.Add(typedef);
                    }
                    break;
                    case DoxMemberKind.Enum:
                    {
                        var enumDoc = new EnumDoc()
                        {
                            Name = member.Name,
                            Descriptions = ParseDescriptions(member),
                        };
                        enumDoc.Values.AddRange(member.EnumValues.Select(x => new EnumValueDoc()
                        {
                            Name = x.Name,
                            Initializer = LinkedTextToString(x.Initializer),
                            Descriptions = ParseDescriptions(x),
                        }));
                        group.Enums.Add(enumDoc);
                    }
                    break;
                    case DoxMemberKind.Variable:
                    {
                        Merge(group.GlobalVariables, ParseVariable(member));
                    }
                    break;
                    default:
                        logger.Warning($"Unknown DoxMemberKinnd '{member.Kind}' for member {member.Name}");
                        break;
                }
            }

            return group;
        }

        private static IEnumerable<ParameterDoc> ParseParameters(MemberDef member)
        {
            foreach (var param in member.Params)
            {
                // If there are no parameters whatsoever, we get a single <param> element
                if (param.DeclName == null && param.DefName == null && param.Type == null)
                    continue;
                // Function declarations may contain a single 'void', which we want to ignore
                if (param.Type?.Type is { Count: 1 } l && l[0] as string == "void")
                    continue;

                string name = param.DeclName ?? param.DefName ?? "";

                // Find its docs...
                var paramDesc = member.DetailedDescription?.Para.SelectMany(x => x.ParameterLists)
                    .Where(x => x.Kind == DoxParamListKind.Param)
                    .SelectMany(x => x.ParameterItems)
                    .SelectMany(x => x.ParameterNameList
                        .Select(name => new { name, desc = x.ParameterDescription }))
                    .FirstOrDefault(x => DocParamNameToString(x.name.ParameterName) == name);

                var direction = paramDesc?.name.ParameterName.Direction switch
                {
                    DoxParamDir.In => ParameterDirection.In,
                    DoxParamDir.Out => ParameterDirection.Out,
                    DoxParamDir.InOut => ParameterDirection.InOut,
                    null or DoxParamDir.None => ParameterDirection.None,
                };

                var functionParameter = new ParameterDoc()
                {
                    Name = name,
                    Type = LinkedTextToString(param.Type),
                    Description = ParasToParagraph(paramDesc?.desc.Para),
                    Direction = direction,
                };

                yield return functionParameter;
            }
        }

        private ClassDoc? ParseInnerClass(string refId)
        {
            var compoundDef = this.ParseDoxygenFile(refId);
            logger.Info($"Parsing class {compoundDef.CompoundName}");

            if (compoundDef.Kind is not (CompoundKind.Struct or CompoundKind.Union))
            {
                logger.Warning($"Don't know how to parse class kind {compoundDef.Kind} in {compoundDef.CompoundName}. Ignoring");
                return null;
            }

            var type = compoundDef.Kind switch
            {
                CompoundKind.Struct => ClassType.Struct,
                CompoundKind.Union => ClassType.Union,
                _ => throw new Exception("Impossible"),
            };

            var cls = new ClassDoc()
            {
                Type = type,
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

        private static VariableDoc ParseVariable(MemberDef member)
        {
            var variable = new VariableDoc()
            {
                Id = member.Id,
                Name = member.Name,
                Type = LinkedTextToString(member.Type),
                Definition = member.Definition,
                Descriptions = ParseDescriptions(member),
                Initializer = LinkedTextToString(member.Initializer),
                Bitfield = member.Bitfield,
                ArgsString = member.ArgsString,
            };
            return variable;
        }

        private static IParagraph ParseReturnDescription(MemberDef member)
        {
            return ParaToParagraph(member.DetailedDescription?.Para.SelectMany(x => x.Parts)
                .OfType<DocSimpleSect>()
                .FirstOrDefault(x => x.Kind == DoxSimpleSectKind.Return)?.Para);
        }

        private static IEnumerable<ReturnValueDoc> ParseReturnValues(MemberDef member)
        {
            var items = member.DetailedDescription?.Para.SelectMany(x => x.ParameterLists)
                .FirstOrDefault(x => x.Kind == DoxParamListKind.RetVal)?.ParameterItems;

            if (items == null)
                yield break;

            foreach (var item in items)
            {
                yield return new ReturnValueDoc()
                {
                    Name = DocParamNameToString(item.ParameterNameList[0].ParameterName) ?? "",
                    Description = ParasToParagraph(item.ParameterDescription.Para),
                };
            }
        }

        private static ReturnDescriptions ParseReturnDescriptions(MemberDef member)
        {
            var descriptions = new ReturnDescriptions()
            {
                Description = ParseReturnDescription(member),
            };
            descriptions.Values.AddRange(ParseReturnValues(member));
            return descriptions;
        }

        private static string? LinkedTextToString(LinkedText? linkedText) =>
            EmbeddedRefTextToString(linkedText?.Type);

        private static string? DocParamNameToString(DocParamName? docParamName) =>
            EmbeddedRefTextToString(docParamName?.Name);

        private static string? EmbeddedRefTextToString(IEnumerable<object>? input)
        {
            if (input == null)
                return null;

            return string.Concat(input.Select(x =>
                x switch
                {
                    string s => s,
                    RefText r => r.Name,
                    _ => throw new ParserException($"Unknown element in DocParamName: {x}"),
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

        private static IParagraph ParaToParagraph(DocPara? para)
        {
            return ParaParser.Parse(para).FirstOrDefault() ?? new TextParagraph();
        }

        private static IParagraph ParasToParagraph(IEnumerable<DocPara>? paras)
        {
            return ParasToParagraphs(paras).FirstOrDefault() ?? new TextParagraph();
        }

        private static IEnumerable<IParagraph> ParasToParagraphs(IEnumerable<DocPara>? paras)
        {
            if (paras == null)
                return Enumerable.Empty<TextParagraph>();

            return paras.SelectMany(x => ParaParser.Parse(x)).Where(x => !x.IsEmpty);
        }

        private static void Merge<T>(List<T> collection, T newItem) where T : IMergable<T>
        {
            var existing = collection.FirstOrDefault(x => x.Id == newItem.Id);
            if (existing != null)
            {
                existing.MergeWith(newItem);
            }
            else
            {
                collection.Add(newItem);
            }
        }

        private CompoundDef ParseDoxygenFile(string refId)
        {
            string filePath = Path.Combine(this.basePath, refId + ".xml");
            var file = Parse<Doxygen>(filePath);
            if (file.CompoundDefs.Count > 1)
            {
                logger.Warning($"File {filePath}: expected 1 compoundDef, got {file.CompoundDefs.Count}. Ignoring all but the first");
            }
            else if (file.CompoundDefs.Count == 0)
            {
                throw new ParserException($"File {filePath} contained 0 compoundDefs");
            }
            return file.CompoundDefs[0];
        }

        private static class SerializerCache<T>
        {
            public static readonly XmlSerializer Instance = new(typeof(T));
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
