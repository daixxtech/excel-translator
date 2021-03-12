using CommandLine;

namespace ExcelTranslator {
    public sealed class Options {
        [Option('e', "excel", Required = true, HelpText = "input excel file path or excel directory path.")]
        public string ExcelPath { get; set; }

        [Option('j', "json", Required = true, HelpText = "export json file path.")]
        public string JSONPath { get; set; }

        [Option('c', "csharp", Required = true, HelpText = "export C# code file path.")]
        public string CSharpCodePath { get; set; }

        [Option('x', "exclude_prefix", Required = false, Default = "", HelpText = "exclude sheet or column start with specified prefix.")]
        public string ExcludePrefix { get; set; }

        [Option('n', "class_namespace", Required = false, Default = "Config", HelpText = "generate class in specified namespace.")]
        public string ClassNamespace { get; set; }

        [Option('p', "class_name_prefix", Required = false, Default = "", HelpText = "generate class name with specified prefix.")]
        public string ClassNamePrefix { get; set; }
    }
}
