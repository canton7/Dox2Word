using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using Dox2Word.Logging;
using Dox2Word.Model;

namespace Dox2Word.Generator
{
    public class DotRenderer
    {
        private static readonly Logger logger = Logger.Instance;

        private readonly string dotPath = "dot.exe";
        private bool haveDot = true;

        public DotRenderer(Dictionary<string, ProjectOption> projectOptions)
        {
            if (projectOptions.TryGetValue("HAVE_DOT", out var haveDot) && Equals(haveDot.TypedValue, false))
            {
                this.haveDot = false;
            }
            if (projectOptions.TryGetValue("DOT_PATH", out var dotPathOption)
                && dotPathOption.TypedValue is string dotPath
                && !string.IsNullOrWhiteSpace(dotPath))
            {
                this.dotPath = dotPath;
            }
        }

        public byte[]? TryRender(string input)
        {
            if (!this.haveDot)
            {
                return null;
            }

            logger.Debug("Rendering dot diagram");

            try
            {
                using var process = new Process()
                {
                    StartInfo = new ProcessStartInfo()
                    {
                        FileName = this.dotPath,
                        Arguments = "-Tpng",
                        UseShellExecute = false,
                        RedirectStandardInput = true,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                    }
                };

                process.Start();

                process.StandardInput.Write(input);
                process.StandardInput.Close();

                var output = new MemoryStream();
                process.StandardOutput.BaseStream.CopyTo(output);

                process.WaitForExit();

                if (process.ExitCode != 0)
                {
                    logger.Warning($"dot.exe exited with code {process.ExitCode}: {process.StandardError.ReadToEnd()}");
                    return null;
                }

                return output.ToArray();
            }
            catch (Win32Exception e) when (e.HResult == -2147467259)
            {
                logger.Warning("Could not find dot.exe. Make sure Graphviz is installed. If you have set DOT_PATH, make sure that it is correct; if not, make sure that dot.exe is in your PATH");
                this.haveDot = false;
                return null;
            }
            catch (Exception e)
            {
                logger.Warning($"Failed to invoke dot.exe: {e.Message}");
                return null;
            }
        }
    }
}
