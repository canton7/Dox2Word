using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Dox2Word.Logging;

namespace Dox2Word.Generator
{
    public class DotRenderer
    {
        private static readonly Logger logger = Logger.Instance;

        private bool haveDot = true;

        public byte[]? TryRender(string input)
        {
            if (!this.haveDot)
            {
                return null;
            }

            logger.Info("Rendering dot diagram");

            try
            {
                using var process = new Process()
                {
                    StartInfo = new ProcessStartInfo()
                    {
                        FileName = "dot.exe",
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
                logger.Warning("Could not find dot.exe in PATH. Make sure Graphviz is installed and added to your PATH");
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
