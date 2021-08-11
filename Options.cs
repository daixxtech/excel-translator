using CommandLine;

namespace ExcelTranslator {
    public sealed class Options {
        [Option('e', "excel", Required = true, HelpText = "source excel file path or excel directory path.")]
        public string ExcelPath { get; set; }

        [Option('j', "json", Required = true, HelpText = "generate json file path.")]
        public string JSONPath { get; set; }

        [Option('c', "csharp", Required = true, HelpText = "generate C# code file path.")]
        public string CSharpCodePath { get; set; }

        [Option('C', "class_prefix", Required = false, HelpText = "generate class name with specified prefix.")]
        public string ClassNamePrefix { get; set; }

        [Option('E', "enum_prefix", Required = false, HelpText = "generate enum name with specified prefix.")]
        public string EnumNamePrefix { get; set; }

        [Option('P', "param_prefix", Required = false, HelpText = "generate param name with specified prefix.")]
        public string ParamNamePrefix { get; set; }

        [Option('n', "namespace", Required = false, Default = "Config", HelpText = "generate code in specified namespace. (only csharp)")]
        public string Namespace { get; set; }
    }
}
