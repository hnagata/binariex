﻿using GlobExpressions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using YamlDotNet.Serialization;

namespace binariex
{
    class BinariexApp
    {
        const string SETTINGS_NAME = "settings.yml";

        static readonly NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();

        readonly IEnumerable<string> inputPaths;
        readonly string settingsPath;

        dynamic settings;
        Dictionary<Glob, string> schemaMap = new Dictionary<Glob, string>();

        public BinariexApp(IEnumerable<string> inputPaths)
        {
            this.inputPaths = inputPaths;
        }

        public BinariexApp(IEnumerable<string> inputPaths, string settingsPath)
        {
            this.inputPaths = inputPaths;
            this.settingsPath = settingsPath;
        }

        public void Run()
        {
            try
            {
                LoadSettings();
                foreach (var path in this.inputPaths)
                {
                    RunWithSingleEntry(path);
                }
            }
            catch (BinariexException exc)
            {
                WriteErrorLog(exc);
            }
            catch (Exception exc)
            {
                logger.Error("Unexpected error occurred. Notify the developers.");
                logger.Debug(exc);
            }
        }

        void LoadSettings()
        {
            var exeDir = Directory.GetParent(Assembly.GetExecutingAssembly().Location).FullName;
            var settingsPath = this.settingsPath ?? Path.Combine(exeDir, SETTINGS_NAME);
            if (!File.Exists(settingsPath))
            {
                throw new BinariexException("loading setting file", "Setting file not found.");
            }

            try
            {
                var deserializer = new Deserializer();
                using (var reader = new StreamReader(File.OpenRead(settingsPath)))
                {
                    this.settings = deserializer.Deserialize<dynamic>(reader);
                }
            }
            catch (Exception exc)
            {
                if (exc is YamlDotNet.Core.SyntaxErrorException || exc is YamlDotNet.Core.SemanticErrorException)
                {
                    throw new BinariexException(exc, "loading setting file", "Invalid syntax/semantics within the setting file.")
                        .AddInputPath($@"{Path.GetFileName(settingsPath)} {Regex.Match(exc.Message, @"\(Line:.+?\)").Value}");
                }
                else
                {
                    throw exc;
                }
            }

            if (!(this.settings is Dictionary<object, object> && (this.settings as Dictionary<object, object>).ContainsKey("schemaMapping")))
            {
                throw new BinariexException("loading setting file", "No \"schemaMapping\" section in the setting file.");
            }
            var schemaMapSection = this.settings["schemaMapping"] as Dictionary<dynamic, dynamic>;
            foreach (var entry in schemaMapSection)
            {
                this.schemaMap.Add(new Glob(entry.Key), entry.Value);
            }
        }

        void RunWithSingleEntry(string inputPath)
        {
            if (Directory.Exists(inputPath))
            {
                foreach (var fileName in Directory.EnumerateFiles(inputPath))
                {
                    RunWithSingleFileSafe(Path.Combine(inputPath, fileName));
                }
            }
            else if (File.Exists(inputPath))
            {
                RunWithSingleFileSafe(inputPath);
            }
            else
            {
                var specialCharPos = inputPath.IndexOfAny("*?[{".ToCharArray());
                if (specialCharPos < 0)
                {
                    throw new BinariexException("finding source file", "Input file not found.").AddInputPath(inputPath);
                }

                var sepCharPos = inputPath.LastIndexOf(Path.DirectorySeparatorChar, specialCharPos);
                if (sepCharPos < 0)
                {
                    throw new BinariexException("finding source file", "Need directory path before wildcard.").AddInputPath(inputPath);
                }

                var rootDirPath = inputPath.Substring(0, sepCharPos);
                if (!Directory.Exists(rootDirPath))
                {
                    throw new BinariexException("finding source file", "Input file not found.").AddInputPath(inputPath);
                }

                foreach (var file in Glob.Files(new DirectoryInfo(rootDirPath), inputPath.Substring(sepCharPos + 1)))
                {
                    RunWithSingleFileSafe(file.FullName);
                }
            }
        }

        void RunWithSingleFileSafe(string inputPath)
        {
            try
            {
                logger.Info("Start converting...");
                logger.Info("  Path: {0}", inputPath);

                RunWithSingleFile(inputPath);

                logger.Info("Finished.");
            }
            catch (BinariexException exc)
            {
                WriteErrorLog(exc);
                logger.Warn("Skipped. ({0})", Path.GetFileName(inputPath));
            }
        }

        void RunWithSingleFile(string inputPath)
        {
            var schemaPath = SelectSchema(inputPath);
            if (!File.Exists(schemaPath))
            {
                throw new BinariexException("finding schema file", "Schema file not found.").AddSchemaPath(schemaPath);
            }

            var outputPathRepr = this.settings["outputPath"] as string;
            var outputPath = Regex.Replace(outputPathRepr, @"\{.+?\}", m =>
                m.Value == "{sourceDirPath}" ? Path.GetDirectoryName(inputPath) :
                m.Value == "{sourceName}" ? Path.GetFileName(inputPath) :
                m.Value == "{sourceBaseName}" ? Path.GetFileNameWithoutExtension(inputPath) :
                m.Value == "{sourceExtension}" ? Path.GetExtension(inputPath) :
                m.Value == "{targetExtension}" ? inputPath.EndsWith(".xlsx") ? ".bin" : ".xlsx" :
                m.Value
            );
            if (File.Exists(outputPath))
            {
                try
                {
                    using (var tmpStream = File.Open(outputPath, FileMode.Open, FileAccess.Read, FileShare.Write))
                    {
                    }
                }
                catch (IOException exc)
                {
                    throw new BinariexException(exc, "writing to file", "Output file is locked.").AddOutputPath(outputPath);
                }
            }

            try
            {
                RunWithSingleFile(inputPath, outputPath, schemaPath);
            }
            catch (BinariexException exc)
            {
                throw exc.AddInputPath(inputPath).AddOutputPath(outputPath).AddSchemaPath(schemaPath);
            }
        }

        void RunWithSingleFile(string inputPath, string outputPath, string schemaPath)
        {
            var schemaDoc = LoadSchemaDoc(schemaPath);

            if (inputPath.EndsWith(".xlsx"))
            {
                using (var reader = new ExcelReader(inputPath))
                {
                    using (var writer = new BinaryWriter(outputPath))
                    {
                        using (var converter = new Converter(reader, writer, schemaDoc))
                        {
                            converter.Run();
                        }
                    }
                }
            }
            else
            {
                using (var reader = new BinaryReader(inputPath))
                {
                    using (var writer = new ExcelWriter(outputPath, this.settings["excel"]))
                    {
                        using (var converter = new Converter(reader, writer, schemaDoc))
                        {
                            converter.Run();
                        }
                        writer.Save();
                    }
                }
            }
        }

        string SelectSchema(string path)
        {
            foreach (var entry in this.schemaMap)
            {
                if (entry.Key.IsMatch(path))
                {
                    return entry.Value;
                }
            }
            throw new BinariexException("finding schema file", "Appropriate schema file not found.").AddInputPath(path);
        }

        XDocument LoadSchemaDoc(string path)
        {
            try
            {
                return XDocument.Load(path, LoadOptions.SetLineInfo);
            }
            catch (Exception exc)
            {
                throw new BinariexException(exc, "loading schema file");
            }
        }

        void WriteErrorLog(BinariexException exc)
        {
            logger.Error("Error occured while {0}:", exc.Stage);
            logger.Error(exc.Message ?? exc.InnerException.Message, exc.MessageParams);
            if (exc.InputPath != null)
            {
                if (exc.ReaderPosition != null)
                {
                    logger.Error("  Input:  {0} ({1})", exc.InputPath, exc.ReaderPosition);
                }
                else
                {
                    logger.Error("  Input:  {0}", exc.InputPath);
                }
            }
            if (exc.OutputPath != null)
            {
                if (exc.WriterPosition != null)
                {
                    logger.Error("  Output: {0} ({1})", exc.OutputPath, exc.WriterPosition);
                }
                else
                {
                    logger.Error("  Output: {0}", exc.OutputPath);
                }
            }
            if (exc.SchemaLineInfo != null)
            {
                logger.Error(
                    "  Schema: {0} (ln.{1}:{2})",
                    exc.SchemaPath, exc.SchemaLineInfo.LineNumber, exc.SchemaLineInfo.LinePosition
                );
            }
            if (exc.InnerException != null && exc.InnerException.Message != exc.Message)
            {
                logger.Error("Details:");
                logger.Error("{0}: {1}", exc.InnerException.GetType(), exc.InnerException.Message);
            }
            if (exc.InnerException != null)
            {
                logger.Debug(exc.InnerException);
            }
        }
    }
}
