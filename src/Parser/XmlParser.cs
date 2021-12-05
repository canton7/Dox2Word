using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using Dox2Word.Logging;
using Dox2Word.Model;
using Dox2Word.Parser.Models;

namespace Dox2Word.Parser
{
    public partial class XmlParser
    {
        private static readonly Logger logger = Logger.Instance;

        private readonly string basePath;
        private readonly Index index;
        private readonly MixedModeParser mixedModeParser;
        private readonly Project project;

        private XmlParser(string basePath, DoxygenIndex doxygenIndex)
        {
            this.basePath = basePath;
            this.index = new Index(doxygenIndex);
            this.mixedModeParser = new MixedModeParser(this.basePath, this.index);

            this.project = new Project();
        }

        public static Project Parse(string basePath)
        {
            try
            {
                string indexFile = Path.Combine(basePath, "index.xml");
                var index = Parse<DoxygenIndex>(indexFile);

                return new XmlParser(basePath, index).Parse();
            }
            catch (ParserException e)
            {
                logger.Error(e);
                throw;
            }
        }

        private Project Parse()
        {
            try
            {
                string doxyfileFile = Path.Combine(this.basePath, "Doxyfile.xml");
                if (File.Exists(doxyfileFile))
                {
                    var doxyfile = Parse<DoxygenFile>(doxyfileFile);
                    this.project.Options.AddRange(this.ParseOptions(doxyfile));
                }

                var allFileCompoundDefs = this.index.Compounds.Values.Where(x => x.Kind == CompoundKind.File)
                    .ToDictionary(x => x.RefId, x => this.ParseCompoundDef(x.RefId));

                var allGroupCompoundDefs = this.index.Compounds.Values.Where(x => x.Kind == CompoundKind.Group)
                    .ToDictionary(x => x.RefId, x => this.ParseCompoundDef(x.RefId));
                this.project.AllGroups = allGroupCompoundDefs.Values.ToDictionary(x => x.Id, x => this.ParseGroup(x));

                // Wire up the group->inner group relationships
                logger.Debug("Creating group to inner group relationships");
                var rootGroups = this.project.AllGroups.Values.ToHashSet();
                foreach (var group in this.project.AllGroups.Values)
                {
                    group.SubGroups.AddRange(allGroupCompoundDefs[group.Id].InnerGroups.Select(x => this.project.AllGroups[x.RefId]));
                    rootGroups.ExceptWith(group.SubGroups);
                }

                this.project.RootGroups.AddRange(rootGroups.OrderBy(x => x.Name));

                // Wire up group -> referenced group relationships
                logger.Debug("Creating group to referenced group relationships");
                var fileIdToOwningGroup = this.project.AllGroups.Values
                    .SelectMany(g => g.Files.Select(f => (file: f, group: g)))
                    .ToDictionary(x => x.file.Id, x => x.group);
                foreach (var fileDef in allFileCompoundDefs.Values)
                {
                    foreach (var include in fileDef.Includes.Where(x => x.IsLocal == DoxBool.Yes && x.RefId != null))
                    {
                        var includingGroup = fileIdToOwningGroup[fileDef.Id];
                        var includedGroup = fileIdToOwningGroup[include.RefId!];
                        if (includingGroup != includedGroup)
                        {
                            if (!includingGroup.IncludedGroups.Contains(includedGroup))
                            {
                                includingGroup.IncludedGroups.Add(includedGroup);
                            }
                            if (!includedGroup.IncludingGroups.Contains(includingGroup))
                            {
                                includedGroup.IncludingGroups.Add(includingGroup);
                            }
                        }
                    }
                }
                foreach (var group in this.project.AllGroups.Values)
                {
                    group.IncludedGroups.Sort((x, y) => x.Id.CompareTo(y.Id));
                    group.IncludingGroups.Sort((x, y) => x.Id.CompareTo(y.Id));
                }

                // Wire up function -> function references
                logger.Debug("Creating function to function references");
                foreach (var groupDef in allGroupCompoundDefs.Values)
                {
                    foreach (var function in groupDef.Sections.SelectMany(x => x.Members).Where(x => x.Kind == DoxMemberKind.Function))
                    {
                        if (this.project.AllFunctions.TryGetValue(function.Id, out var functionDoc))
                        {
                            foreach (var references in function.References)
                            {
                                if (this.project.AllFunctions.TryGetValue(references.RefId, out var referencesDoc))
                                {
                                    functionDoc.References.Add(referencesDoc);
                                }
                            }
                            foreach (var referencedBy in function.ReferencedBy)
                            {
                                if (this.project.AllFunctions.TryGetValue(referencedBy.RefId, out var referencedByDoc))
                                {
                                    functionDoc.ReferencedBy.Add(referencedByDoc);
                                }
                            }
                        }
                    }
                }
            }
            catch (ParserException e)
            {
                logger.Error(e);
            }

            return this.project;
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

        private Group ParseGroup(CompoundDef compoundDef)
        {
            logger.Debug($"Parsing compound {compoundDef.CompoundName}");

            var group = new Group()
            {
                Id = compoundDef.Id,
                Name = compoundDef.Title,
                Descriptions = this.ParseDescriptions(compoundDef),
            };
            group.Files.AddRange(compoundDef.InnerFiles.Select(x => new FileDoc() { Id = x.RefId, Name = x.Name }));
            group.Classes.AddRange(compoundDef.InnerClasses.Select(x => this.ParseInnerClass(x.RefId)).Where(x => x != null)!);

            var members = compoundDef.Sections.SelectMany(x => x.Members);
            foreach (var member in members)
            {
                // If they didn't document it, don't include it. Doxygen itself will warn if something should have been documented
                // but wasn't.
                if (member.BriefDescription?.Para.Count is 0 or null && member.DetailedDescription?.Para.Count is 0 or null)
                    continue;

                logger.Debug($"Parsing member {member.Name}");
                switch (member.Kind)
                {
                    case DoxMemberKind.Function:
                    {
                        var function = new FunctionDoc()
                        {
                            Id = member.Id,
                            Name = member.Name,
                            Descriptions = this.ParseDescriptions(member),
                            ReturnType = this.LinkedTextToRuns(member.Type),
                            ReturnDescriptions = this.ParseReturnDescriptions(member),
                            Definition = member.Definition ?? "",
                            ArgsString = member.ArgsString ?? "",
                        };
                        function.Parameters.AddRange(this.ParseParameters(member));
                        this.project.AllFunctions.Add(member.Id, function);
                        group.Functions.Add(function);
                    }
                    break;
                    case DoxMemberKind.Define:
                    {
                        var macro = new MacroDoc()
                        {
                            Id = member.Id,
                            Name = member.Name,
                            Descriptions = this.ParseDescriptions(member),
                            ReturnDescriptions = this.ParseReturnDescriptions(member),
                            Initializer = this.LinkedTextToRuns(member.Initializer),
                            HasParameters = member.Params.Count > 0,
                        };
                        macro.Parameters.AddRange(this.ParseParameters(member));
                        group.Macros.Add(macro);
                    }
                    break;
                    case DoxMemberKind.Typedef:
                    {
                        var typedef = new TypedefDoc()
                        {
                            Id = member.Id,
                            Name = member.Name,
                            Type = this.LinkedTextToRuns(member.Type),
                            Definition = member.Definition ?? "",
                            Descriptions = this.ParseDescriptions(member),
                        };
                        group.Typedefs.Add(typedef);
                    }
                    break;
                    case DoxMemberKind.Enum:
                    {
                        var enumDoc = new EnumDoc()
                        {
                            Id = member.Id,
                            Name = member.Name,
                            Descriptions = this.ParseDescriptions(member),
                        };
                        enumDoc.Values.AddRange(member.EnumValues.Select(x => new EnumValueDoc()
                        {
                            Id = x.Id,
                            Name = x.Name,
                            Initializer = this.LinkedTextToRuns(x.Initializer),
                            Descriptions = this.ParseDescriptions(x),
                        }));
                        group.Enums.Add(enumDoc);
                    }
                    break;
                    case DoxMemberKind.Variable:
                    {
                        Merge(group.GlobalVariables, this.ParseVariable(member));
                    }
                    break;
                    default:
                        logger.Unsupported($"Unknown DoxMemberKind '{member.Kind}' for member {member.Name}. Ignoring");
                        break;
                }
            }

            return group;
        }

        private IEnumerable<ParameterDoc> ParseParameters(MemberDef member)
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
                    Type = this.LinkedTextToRuns(param.Type),
                    Direction = direction,
                };
                functionParameter.Description.AddRange(this.ParasToParagraphs(paramDesc?.desc.Para));

                yield return functionParameter;
            }
        }

        private ClassDoc? ParseInnerClass(string refId)
        {
            var compoundDef = this.ParseCompoundDef(refId);
            logger.Debug($"Parsing class {compoundDef.CompoundName}");

            if (compoundDef.Kind is not (CompoundKind.Struct or CompoundKind.Union))
            {
                logger.Unsupported($"Don't know how to parse class kind {compoundDef.Kind} in {compoundDef.CompoundName}. Ignoring");
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
                Id = compoundDef.Id,
                Name = compoundDef.CompoundName ?? "",
                Descriptions = this.ParseDescriptions(compoundDef),
            };

            var members = compoundDef.Sections.SelectMany(x => x.Members)
                .Where(x => x.Kind == DoxMemberKind.Variable);
            foreach (var member in members)
            {
                cls.Variables.Add(this.ParseVariable(member));
            }

            return cls;
        }

        private VariableDoc ParseVariable(MemberDef member)
        {
            var variable = new VariableDoc()
            {
                Id = member.Id,
                Name = member.Name,
                Type = this.LinkedTextToRuns(member.Type),
                Definition = member.Definition,
                Descriptions = this.ParseDescriptions(member),
                Initializer = this.LinkedTextToRuns(member.Initializer),
                Bitfield = member.Bitfield,
                ArgsString = member.ArgsString,
            };
            return variable;
        }

        private IEnumerable<IParagraph> ParseReturnDescription(MemberDef member)
        {
            return this.ParaToParagraphs(member.DetailedDescription?.Para.SelectMany(x => x.Parts)
                .OfType<DocSimpleSect>()
                .FirstOrDefault(x => x.Kind == DoxSimpleSectKind.Return)?.Para);
        }

        private IEnumerable<ReturnValueDoc> ParseReturnValues(MemberDef member)
        {
            var items = member.DetailedDescription?.Para.SelectMany(x => x.ParameterLists)
                .FirstOrDefault(x => x.Kind == DoxParamListKind.RetVal)?.ParameterItems;

            if (items == null)
                yield break;

            foreach (var item in items)
            {
                var doc = new ReturnValueDoc()
                {
                    Name = DocParamNameToString(item.ParameterNameList[0].ParameterName) ?? "",
                };
                doc.Description.AddRange(this.ParasToParagraphs(item.ParameterDescription.Para));
                yield return doc;
            }
        }

        private ReturnDescriptions ParseReturnDescriptions(MemberDef member)
        {
            var descriptions = new ReturnDescriptions();
            descriptions.Description.AddRange(this.ParseReturnDescription(member));
            descriptions.Values.AddRange(this.ParseReturnValues(member));
            return descriptions;
        }

        private List<TextRun> LinkedTextToRuns(LinkedText? linkedText) =>
            this.EmbeddedRefTextToRuns(linkedText?.Type).ToList();

        private static string? DocParamNameToString(DocParamName? docParamName) =>
            EmbeddedRefTextToString(docParamName?.Name);

        private IEnumerable<TextRun> EmbeddedRefTextToRuns(IEnumerable<object>? input)
        {
            if (input == null)
                yield break;

            foreach (object element in input)
            {
                yield return element switch
                {
                    string s => new TextRun(s),
                    RefText r => new TextRun(r.Name, referenceId: this.index.ShouldReference(r.RefId) ? r.RefId : null),
                    _ => throw new ParserException($"Unknown element in EmbeddedRefText: {element}"),
                };
            }
        }

        private static string? EmbeddedRefTextToString(IEnumerable<object>? input)
        {
            if (input == null)
                return null;

            return string.Concat(input.Select(x =>
                x switch
                {
                    string s => s,
                    RefText r => r.Name,
                    _ => throw new ParserException($"Unknown element in EmbeddedRefText: {x}"),
                }));
        }

        private Descriptions ParseDescriptions(IDoxDescribable member)
        {
            var descriptions = new Descriptions()
            {
                BriefDescription = this.ParasToParagraph(member.BriefDescription?.Para),
            };
            descriptions.DetailedDescription.AddRange(this.ParasToParagraphs(member.DetailedDescription?.Para));
            return descriptions;
        }

        private IEnumerable<IParagraph> ParaToParagraphs(DocPara? para)
        {
            return this.mixedModeParser.Parse(para?.Parts).Where(x => !x.IsEmpty);
        }

        private IParagraph ParasToParagraph(IEnumerable<DocPara>? paras)
        {
            return this.ParasToParagraphs(paras).FirstOrDefault() ?? new TextParagraph();
        }

        private IEnumerable<IParagraph> ParasToParagraphs(IEnumerable<DocPara>? paras)
        {
            if (paras == null)
                return Enumerable.Empty<TextParagraph>();

            return paras.SelectMany(x => this.mixedModeParser.Parse(x.Parts)).Where(x => !x.IsEmpty);
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

        private CompoundDef ParseCompoundDef(string refId)
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
        private static T Parse<T>(string filePath)
        {
            using (var stream = File.OpenRead(filePath))
            using (var reader = new XmlTextReader(stream))
            {
                reader.WhitespaceHandling = WhitespaceHandling.All;
                return (T)SerializerCache.Get<T>().Deserialize(reader)!;
            }
        }
    }
}
