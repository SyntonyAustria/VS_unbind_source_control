// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
// <copyright file="Program.cs" company="Syntony">
//     Copyright © 2016-2016 by Syntony - http://members.aon.at/hahnl - All rights reserved.
// </copyright>
// <author>Ing. Josef Hahnl, MBA - User</author>
// <email>hahnl@aon.at</email>
// <date>24.08.2016 10:51:33</date>
// <information solution="VSUnbindSourceControl" project="VSUnbindSourceControl" framework=".NET Framework 4.0" kind="Windows (C#)">
//     <file type=".cs" created="17.08.2016 17:33:10" modified="24.08.2016 10:51:33" lastAccess="24.08.2016 10:51:33">
//         E:\Syntony\Projects\Tools.Collection\VS_unbind_source_control\VSUnbindSourceControl\Program.cs
//     </file>
//     <lines total="413" netLines="368" codeLines="285" allCommentLines="83" commentLines="49" documentationLines="34" blankLines="45" codeRatio="69.01 %"/>
//     <language>C#</language>
//     <identifiers>
//         <namespace>VSUnbindSourceControl</namespace>
//         <class>Program</class>
//     </identifiers>
// </information>
// <summary>
//     The program.
// </summary>
// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
namespace VSUnbindSourceControl
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Diagnostics;
    using System.Diagnostics.CodeAnalysis;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using System.Text;
    using System.Xml;
    using System.Xml.Linq;

    using JetBrains.Annotations;

    /// <summary>The program.</summary>
    public class Program
    {
        /// <summary>Deletes the file.</summary>
        /// <param name="filename">The filename.</param>
        public static void DeleteFile([NotNull] string filename)
        {
            File.SetAttributes(filename, FileAttributes.Normal);
            File.Delete(filename);
        }

        /// <summary>Main program.</summary>
        /// <param name="args">The arguments.</param>
        public static void Main([CanBeNull] string[] args)
        {
            string folder;
            if (args == null || args.Length < 1)
            {
                Assembly entryAssembly = Assembly.GetEntryAssembly();
                folder = Path.GetDirectoryName(entryAssembly.Location);
            }
            else
            {
                folder = args[0].Trim();
            }

            if (folder != null && folder.Length < 1)
            {
                WriteLine(ConsoleColor.Red, "ERROR: empty folder name");
                WriteLine(ConsoleColor.Red, "Stopping ...");
                Console.ReadLine();
                return;
            }

            if (!Directory.Exists(folder))
            {
                WriteLine(ConsoleColor.Red, "ERROR: Folder does not exist");
                WriteLine(ConsoleColor.Red, "Stopping ...");
                Console.ReadLine();
                return;
            }

            Console.WriteLine("Starting ...");

            List<string> sccFilesToDelete = new List<string>();
            List<string> projFilesToModify = new List<string>();
            List<string> slnFilesToModify = new List<string>();
            List<string> files = new List<string>(Directory.GetFiles(folder, "*.*", SearchOption.AllDirectories));

            foreach (string filename in files)
            {
                string normalizedFilename = filename.ToLower();
                if (normalizedFilename.Contains(".") &&
                    normalizedFilename.EndsWith("proj", StringComparison.OrdinalIgnoreCase) &&
                    !normalizedFilename.EndsWith("vdproj", StringComparison.OrdinalIgnoreCase))
                {
                    projFilesToModify.Add(filename);
                }
                else if (normalizedFilename.EndsWith(".sln", StringComparison.OrdinalIgnoreCase))
                {
                    slnFilesToModify.Add(filename);
                }
                else if (normalizedFilename.EndsWith(".vssscc", StringComparison.OrdinalIgnoreCase) ||
                         normalizedFilename.EndsWith(".vspscc", StringComparison.OrdinalIgnoreCase))
                {
                    sccFilesToDelete.Add(filename);
                }
                else
                {
                    // do nothing
                }
            }

            if (projFilesToModify.Count + slnFilesToModify.Count + sccFilesToDelete.Count < 1)
            {
                Console.WriteLine("No files to modify or delete. Exiting.");
                Console.ReadLine();
                return;
            }

            ProcessFile(ModifySolutionFile, slnFilesToModify);
            ProcessFile(ModifyProjectFile, projFilesToModify);
            ProcessFile(DeleteFile, sccFilesToDelete);
            DirectoryRecursiveSearch(folder,
                                     path =>
                                     {
                                         if (!path.EndsWith("$tf", StringComparison.OrdinalIgnoreCase) && !path.EndsWith(".git", StringComparison.OrdinalIgnoreCase))
                                         {
                                             return;
                                         }

                                         try
                                         {
                                             Directory.Delete(path, true);
                                             WriteLine(ConsoleColor.Green, $"Directory : {path} deleted.");
                                         }
                                         catch (UnauthorizedAccessException ex)
                                         {
                                             RestartAsAdministrator();
                                             WriteLine(ConsoleColor.Red, ex.GetType() + " " + ex.Message);
                                         }
                                         catch (Exception ex)
                                         {
                                             WriteLine(ConsoleColor.Red, ex.GetType() + " " + ex.Message);
                                         }
                                     });

            WriteLine(ConsoleColor.Green, "Done ...");
            WriteLine(ConsoleColor.White, "Press ENTER to exit ...");
            Console.ReadLine();
        }

        /// <summary>Modifies the project file.</summary>
        /// <param name="filename">The filename.</param>
        /// <exception cref="ArgumentException">Internal Error: ModifyProjectFile called with a file that is not a project</exception>
        public static void ModifyProjectFile([NotNull] string filename)
        {
            if (!filename.ToLower().EndsWith("proj", StringComparison.Ordinal))
            {
                throw new ArgumentException("Internal Error: ModifyProjectFile called with a file that is not a project");
            }

            Console.WriteLine($"Modifying Project : {filename}");

            // Load the Project file
            XDocument doc;
            Encoding encoding = new UTF8Encoding(false);
            using (StreamReader reader = new StreamReader(filename, encoding))
            {
                doc = XDocument.Load(reader);
                encoding = reader.CurrentEncoding;
            }

            // Modify the Source Control Elements
            RemoveSccElementsAttributes(doc.Root);

            // Remove the read-only flag
            FileAttributes originalAttr = File.GetAttributes(filename);
            File.SetAttributes(filename, FileAttributes.Normal);

            // if the original document doesn't include the encoding attribute
            // in the declaration then do not write it to the outpu file.
            if (string.IsNullOrEmpty(doc.Declaration?.Encoding))
            {
                encoding = null;
            }
            else if (!doc.Declaration.Encoding.StartsWith("utf", StringComparison.OrdinalIgnoreCase))
            {
                // else if its not utf (i.e. utf-8, utf-16, utf32) format which use a BOM
                // then use the encoding identified in the XML file.
                encoding = Encoding.GetEncoding(doc.Declaration.Encoding);
            }

            // Write out the XML
            using (XmlTextWriter writer = new XmlTextWriter(filename, encoding))
            {
                writer.Formatting = Formatting.Indented;
                doc.Save(writer);
                writer.Close();
            }

            // Restore the original file attributes
            File.SetAttributes(filename, originalAttr);
        }

        /// <summary>Modifies the solution file.</summary>
        /// <param name="filename">The filename.</param>
        /// <exception cref="ArgumentException">Internal Error: ModifySolutionFile called with a file that is not a solution</exception>
        public static void ModifySolutionFile([NotNull] string filename)
        {
            if (!filename.ToLower().EndsWith(".sln", StringComparison.Ordinal))
            {
                throw new ArgumentException("Internal Error: ModifySolutionFile called with a file that is not a solution");
            }

            Console.WriteLine("Modifying Solution: {0}", filename);

            // Remove the read-only flag
            FileAttributes originalAttr = File.GetAttributes(filename);
            File.SetAttributes(filename, FileAttributes.Normal);

            List<string> outputLines = new List<string>();

            bool inSourcecontrolSection = false;

            Encoding encoding;
            string[] lines = ReadAllLines(filename, out encoding);

            foreach (string line in lines)
            {
                string lineTrimmed = line.Trim();

                // lines can contain separators which interferes with the regex
                // escape them to prevent regex from having problems
                lineTrimmed = Uri.EscapeDataString(lineTrimmed);

                if (lineTrimmed.StartsWith("GlobalSection(SourceCodeControl)", StringComparison.Ordinal)
                    || lineTrimmed.StartsWith("GlobalSection(TeamFoundationVersionControl)", StringComparison.Ordinal)
                    || System.Text.RegularExpressions.Regex.IsMatch(lineTrimmed, @"GlobalSection\(.*Version.*Control", System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                {
                    // this means we are starting a Source Control Section
                    // do not copy the line to output
                    inSourcecontrolSection = true;
                }
                else if (inSourcecontrolSection && lineTrimmed.StartsWith("EndGlobalSection", StringComparison.Ordinal))
                {
                    // This means we were Source Control section and now see the ending marker
                    // do not copy the line containing the ending marker
                    inSourcecontrolSection = false;
                }
                else if (lineTrimmed.StartsWith("Scc", StringComparison.Ordinal))
                {
                    // These lines should be ignored completely no matter where they are seen
                }
                else
                {
                    // No handle every other line
                    // Basically as long as we are not in a source control section
                    // then that line can be copied to output
                    if (!inSourcecontrolSection)
                    {
                        outputLines.Add(line);
                    }
                }
            }

            // Write the file back out
            File.WriteAllLines(filename, outputLines, encoding);

            // Restore the original file attributes
            File.SetAttributes(filename, originalAttr);
        }

        /// <summary>A recursive directory search.</summary>
        /// <param name="folder">The folder to make a recursive search for.</param>
        /// <param name="actionForEachDirectory">The action to perform on each found folder.</param>
        private static void DirectoryRecursiveSearch([CanBeNull] string folder, [CanBeNull] Action<string> actionForEachDirectory)
        {
            if (string.IsNullOrWhiteSpace(folder) || folder.EndsWith(".vs", StringComparison.OrdinalIgnoreCase) || folder.Equals("$tf", StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            try
            {
                actionForEachDirectory?.Invoke(folder);
                foreach (string subfolder in Directory.GetDirectories(folder))
                {
                    DirectoryRecursiveSearch(new Uri(subfolder).LocalPath, actionForEachDirectory);
                }
            }
            catch (Exception ex)
            {
                WriteLine(ConsoleColor.Red, ex.Message);
            }
        }

        /// <summary>Processes a list of files based on the processing method.</summary>
        /// <param name="processMethod">The method for processing the files.</param>
        /// <param name="files">The list of files.</param>
        private static void ProcessFile([NotNull] Action<string> processMethod, [NotNull] IEnumerable<string> files)
        {
            foreach (string file in files)
            {
                try
                {
                    processMethod(file);
                }
                catch (Exception e)
                {
                    string message = string.Format("Unable to process {0}: {1}", file, e.Message);
                    WriteLine(ConsoleColor.Red, message);
                }
            }
        }

        /// <summary>Reads all the lines from a test file into an array.</summary>
        /// <param name="path">The file to open for reading.</param>
        /// <param name="encoding">The file encoding.</param>
        /// <returns>A string array containing all the lines from the file</returns>
        /// <remarks>
        /// UTF-8 encoded files optionally include a byte order mark (BOM) at the beginning of the file.
        /// If the mark is detected by the StreamReader class, it will modify it's encoding property so that it
        /// reflects that file was written with a BOM. However, if no BOM is detected the StreamReader will not
        /// modify it encoding property. The determined UTF-8 encoding (UTF-8 with BOM or UTF-8 without BOM) is
        /// returned as an output parameter.
        /// </remarks>
        [NotNull]
        private static string[] ReadAllLines([NotNull] string path, [NotNull] out Encoding encoding)
        {
            List<string> lines = new List<string>();

            Encoding encodingNoBom = new UTF8Encoding(false);
            using (StreamReader reader = new StreamReader(path, encodingNoBom))
            {
                while (!reader.EndOfStream)
                {
                    lines.Add(reader.ReadLine());
                }

                encoding = reader.CurrentEncoding;
            }

            return lines.ToArray();
        }

        /// <summary>Removes the SCC elements attributes.</summary>
        /// <param name="element">The element.</param>
        private static void RemoveSccElementsAttributes([CanBeNull] XElement element)
        {
            if (element == null)
            {
                return;
            }

            element.Elements().Where(x => x.Name.LocalName.StartsWith("Scc", StringComparison.Ordinal)).Remove();
            element.Attributes().Where(x => x.Name.LocalName.StartsWith("Scc", StringComparison.Ordinal)).Remove();

            foreach (XElement child in element.Elements().Where(e => e != null))
            {
                RemoveSccElementsAttributes(child);
            }
        }

        /// <summary>Restarts the application as administrator.</summary>
        [SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification = "okay here")]
        private static void RestartAsAdministrator()
        {
            Assembly entryAssembly = Assembly.GetEntryAssembly();
            string fileName = entryAssembly.Location;
            if (string.IsNullOrEmpty(fileName))
            {
                return;
            }

            ProcessStartInfo processStartInfo = new ProcessStartInfo
            {
                UseShellExecute = true,
                WorkingDirectory = Environment.CurrentDirectory,
                FileName = fileName,
                Verb = "runas"
            };
            try
            {
                Process.Start(processStartInfo);
            }
            catch (Win32Exception)
            {
                // An error occurred when opening the associated file.
                // or
                // The sum of the length of the arguments and the length of the full path to the process exceeds 2080.
                // The error message associated with this exception can be one of the following: "The data area passed to a system call is too small." or "Access is denied."
                return;
            }
            catch (Exception ex)
            {
                WriteLine(ConsoleColor.Red, ex.GetType() + " " + ex.Message);
                return;
            }

            Environment.Exit(0);
        }

        /// <summary>Writes a line to console in the specified foreground color.</summary>
        /// <param name="foregroundColor">The foreground color.</param>
        /// <param name="value">The value that is written to the console.</param>
        private static void WriteLine(ConsoleColor foregroundColor, [CanBeNull] string value)
        {
            ConsoleColor current = Console.ForegroundColor;
            Console.ForegroundColor = foregroundColor;
            Console.WriteLine(value ?? "__null__");
            Console.ForegroundColor = current;
        }
    }
}