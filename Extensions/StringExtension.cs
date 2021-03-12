using System;

namespace ExcelTranslator.Extensions {
    public enum NamingStyle {
        // ReSharper disable once InconsistentNaming
        lowerCamel,
        // ReSharper disable once InconsistentNaming
        UpperCamel,
        // ReSharper disable once InconsistentNaming
        all_lower,
        // ReSharper disable once InconsistentNaming
        ALL_UPPER
    }

    public static class StringExtension {
        /// <summary> 转换到对应的命名样式 </summary>
        public static string ToNamingStyle(this string name, NamingStyle style) {
            if (!name.IsValidName()) {
                Console.WriteLine("Invalid name: {0}", name);
                return null;
            }
            char[] nameCharArr = name.ToCharArray();
            int nameCharCount = nameCharArr.Length;
            switch (style) {
            case NamingStyle.lowerCamel:
                for (int j = 0; j < nameCharCount; j++) {
                    if (nameCharArr[j] < 'A' || nameCharArr[j] > 'Z') {
                        break;
                    }
                    nameCharArr[j] = (char) (nameCharArr[j] + 32);
                }
                break;
            case NamingStyle.UpperCamel:
                if (nameCharArr[0] >= 'a' && nameCharArr[0] <= 'z') {
                    nameCharArr[0] = (char) (nameCharArr[0] - 32);
                }
                break;
            case NamingStyle.all_lower: // TODO
                break;
            case NamingStyle.ALL_UPPER: // TODO
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(style), style, null);
            }
            return new string(nameCharArr);
        }

        /// <summary> 是否为合法的命名 </summary>
        public static bool IsValidName(this string name) {
            return !string.IsNullOrEmpty(name) && (name[0] >= 'A' && name[0] <= 'Z' || name[0] >= 'a' && name[0] <= 'z');
        }
    }
}
