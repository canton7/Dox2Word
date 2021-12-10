#!/usr/bin/env dotnet-script

#r "nuget: SimpleTasks, 0.9.4"

using SimpleTasks;
using static SimpleTasks.SimpleTask;

#nullable enable

string dox2wordDir = "src";
string cTestDir = "test/C";

CreateTask("build").Run((string versionOpt, string configurationOpt) =>
{
    var flags = $"--configuration={configurationOpt ?? "Release"} -p:VersionPrefix=\"{versionOpt ?? "0.0.0"}\"";
    Command.Run("dotnet", $"build {flags} \"{dox2wordDir}\"");
    Command.Run("dotnet", $"build -t:ILRepack {flags} \"{dox2wordDir}\"");
});

CreateTask("test").Run((string configurationOpt) =>
{
	Directory.SetCurrentDirectory(cTestDir);
	Command.Run("doxygen");
	Command.Run($"{Path.Combine("../../bin", configurationOpt ?? "Release", "Dox2Word.exe")}",
		$"-i xml -t Template.docx -o Output.docx -v");
});

return InvokeTask(Args);
