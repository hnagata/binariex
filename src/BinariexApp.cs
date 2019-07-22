using GlobExpressions;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using YamlDotNet.Serialization;

namespace binariex
{
    class BinariexApp
    {
        const string SETTINGS_PATH = "settings.yml";

        string[] inputPaths;

        dynamic settings;
        Dictionary<Glob, string> schemaMap = new Dictionary<Glob, string>();

        public BinariexApp(string[] args)
        {
            this.inputPaths = args;

            var deserializer = new Deserializer();
            using (var reader = new StreamReader(File.OpenRead(SETTINGS_PATH)))
            {
                this.settings = deserializer.Deserialize<dynamic>(reader);
            }

            var schemaMapSection = this.settings["schemaMapping"] as Dictionary<dynamic, dynamic>;
            foreach (var entry in schemaMapSection)
            {
                this.schemaMap.Add(new Glob(entry.Key), entry.Value);
            }
        }

        public void Run()
        {
            foreach (var path in this.inputPaths)
            {
                RunWithSingleEntry(path);
            }
        }

        void RunWithSingleEntry(string inputPath)
        {
            if (Directory.Exists(inputPath))
            {
                foreach (var fileName in Directory.EnumerateFiles(inputPath))
                {
                    RunWithSingleFile(Path.Combine(inputPath, fileName));
                }
            }
            else if (File.Exists(inputPath))
            {
                RunWithSingleFile(inputPath);
            }
            else
            {
                throw new FileNotFoundException(null, inputPath);
            }
        }

        void RunWithSingleFile(string inputPath)
        {
            var schemaPath = SelectSchema(inputPath);
            var schemaDoc = XDocument.Load(schemaPath);

            var outputPathRepr = this.settings["outputPath"] as string;
            var outputPath = Regex.Replace(outputPathRepr, @"\{.+?\}", m =>
                m.Value == "{sourceDirPath}" ? Path.GetDirectoryName(inputPath) :
                m.Value == "{sourceName}" ? Path.GetFileName(inputPath) :
                m.Value == "{sourceBaseName}" ? Path.GetFileNameWithoutExtension(inputPath) :
                m.Value == "{sourceExtension}" ? Path.GetExtension(inputPath) :
                m.Value
            );

            if (inputPath.EndsWith(".xlsx"))
            {
                using (var reader = new ExcelReader(inputPath))
                {
                    using (var writer = new BinaryWriter(outputPath.Replace("{targetExtension}", ".bin")))
                    {
                        new Converter(reader, writer, schemaDoc).Run();
                    }
                }
            }
            else
            {
                using (var reader = new BinaryReader(inputPath))
                {
                    using (var writer = new ExcelWriter(outputPath.Replace("{targetExtension}", ".xlsx"), this.settings["excel"]))
                    {
                        new Converter(reader, writer, schemaDoc).Run();
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
            return null;
        }
    }
}
