﻿#define WINDOWS

using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Resources;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Arc.YTSubConverter.Formats;
using Arc.YTSubConverter.Formats.Ass;

namespace Arc.YTSubConverter.CLI
{
    internal static class Program
    {
        // I really do not know whether I can remove this or not, or
        // whether I *should* remove it or not.
        [STAThread]
        internal static void Main(string[] args)
        {
            PreloadResources();

            RunCommandLine(args);
            return;

        }

        private static void PrintUsageString()
        {
            string ScriptName = Path.GetFileName(Environment.GetCommandLineArgs()[0]);
            Console.WriteLine("Usage: " + ScriptName + " [--visual] INPUT OUTPUT");
        }

        private static void RunCommandLine(string[] args)
        {
            AttachConsole(ATTACH_PARENT_PROCESS);

            CommandLineArguments parsedArgs = ParseArguments(args);
            if (parsedArgs == null)
                return;

            if (!File.Exists(parsedArgs.SourceFilePath))
            {
                Console.WriteLine("Specified source file not found");
                return;
            }

            try
            {
                SubtitleDocument sourceDoc = SubtitleDocument.Load(parsedArgs.SourceFilePath);
                SubtitleDocument destinationDoc =
                    Path.GetExtension(parsedArgs.DestinationFilePath).ToLower() switch
                    {
                        ".ass" => parsedArgs.ForVisualization ? new VisualizingAssDocument(sourceDoc) : new AssDocument(sourceDoc),
                        ".srv3" => new YttDocument(sourceDoc),
                        ".ytt" => new YttDocument(sourceDoc),
                        _ => new SrtDocument(sourceDoc)
                    };
                destinationDoc.Save(parsedArgs.DestinationFilePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error occurred: {ex.Message}");
            }
        }

        private static CommandLineArguments ParseArguments(string[] args)
        {
            CommandLineArguments parsedArgs = new CommandLineArguments();

            List<string> filePaths = new List<string>();
            foreach (string arg in args)
            {
                if (arg.StartsWith("-"))
                {
                    switch (arg)
                    {
                        case "--visual":
                            parsedArgs.ForVisualization = true;
                            break;
                    }
                }
                else
                {
                    filePaths.Add(arg);
                }
            }

            if (filePaths.Count == 0)
            {
                PrintUsageString();
                Console.WriteLine("Please specify an INPUT file.");
                return null;
            }

            if (filePaths.Count > 2)
            {
                PrintUsageString();
                Console.WriteLine("Too many file paths specified.");
                return null;
            }

            parsedArgs.SourceFilePath = filePaths[0];
            if (filePaths.Count == 1)
            {
                string destinationExtension =
                    Path.GetExtension(parsedArgs.SourceFilePath).ToLower() switch
                    {
                        ".ass" => ".ytt",
                        ".ytt" => ".reverse.ass",
                        ".srv3" => ".ass",
                        _ => ".srt"
                    };
                parsedArgs.DestinationFilePath = Path.ChangeExtension(parsedArgs.SourceFilePath, destinationExtension);
            }
            else
            {
                parsedArgs.DestinationFilePath = filePaths[1];
            }

            return parsedArgs;
        }

        /// <summary>
        /// Manually load the resources available in the .exe so the ILMerged release build doesn't need satellite assemblies anymore
        /// </summary>
        // Not exactly sure if we still need this since we aren't using
        // ILMerge anymore.
        private static void PreloadResources()
        {
            PreloadResources<YTSubConverter.Resources>(YTSubConverter.Resources.ResourceManager);
        }

        private static void PreloadResources<TResources>(ResourceManager resourceManager)
        {
            Assembly assembly = Assembly.GetEntryAssembly();
            FieldInfo resourceSetsField = typeof(ResourceManager).GetField("_resourceSets", BindingFlags.NonPublic | BindingFlags.Instance);
            Dictionary<string, ResourceSet> resourceSets = (Dictionary<string, ResourceSet>)resourceSetsField.GetValue(resourceManager);

            foreach (string resourceName in assembly.GetManifestResourceNames())
            {
                Match match = Regex.Match(resourceName, Regex.Escape(typeof(TResources).FullName) + @"\.([-\w]+)\.resources$");
                if (!match.Success)
                    continue;

                string culture = match.Groups[1].Value;
                using Stream stream = assembly.GetManifestResourceStream(resourceName);
                ResourceSet resSet = new ResourceSet(stream);
                resourceSets.Add(culture, resSet);
            }
        }

#if WINDOWS
        [DllImport("kernel32.dll")]
        private static extern bool AttachConsole(int dwProcessId);
#else
        private static bool AttachConsole(int dwProcessId) { return true; }
#endif

        private const int ATTACH_PARENT_PROCESS = -1;

        private class CommandLineArguments
        {
            public bool ForVisualization
            {
                get;
                set;
            }

            public string SourceFilePath
            {
                get;
                set;
            }


            public string DestinationFilePath
            {
                get;
                set;
            }
        }
    }
}
