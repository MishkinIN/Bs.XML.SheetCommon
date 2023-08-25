using Bs.XML.SpreadSheet.Resources;
using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;

namespace Bs.XML.SpreadSheet {
    public struct ColumnName {
        public const int MaxColumnNumber = 16384;
        //private static readonly Collection<char> Letters = new Collection<char>(new Char[] {
        //    'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'
        //});
        private static readonly Regex regex = new Regex("^[A-Z]{1,3}$", RegexOptions.Compiled);

        private readonly uint number;

        public ColumnName(string name)
            : this() {
            number = GetNumber(name);
        }
        public ColumnName(uint number) {
            Throw.IfIsTrue(number > MaxColumnNumber,
                () => new ArgumentOutOfRangeException(nameof(number)));
            this.number = number - 1;
        }
        public static string GetName(uint number) {
            Throw.IfIsTrue(number > MaxColumnNumber,
                () => new ArgumentOutOfRangeException(nameof(number)));
            return GetNameInternal(number);
        }

        public static uint GetNumber(string name) {
            //name = name.ToUpper();
            bool nameOutOfRange = String.IsNullOrEmpty(name) ||
                !regex.IsMatch(name) ||
                (name.Length > 2 & name.CompareTo("XFD") > 0);
            Throw.IfIsTrue(nameOutOfRange,
                () => new ArgumentOutOfRangeException(nameof(name), FormattableStringFactory.Create(ExceptionMessage.ColumnNameWrong, name).ToString()));
            return GetNumberInternal(name);
        }
        public uint Number { get { return number + 1; } }
        public string Name { get { return GetName(number + 1); } }
        public static ColumnName operator +(ColumnName cn, int value) {
            var number = cn.number + value;
            Throw.IfIsTrue(number < 0 | number > MaxColumnNumber - 1,
                () => new OverflowException());
            return new ColumnName((uint) number + 1);
    }
    public static ColumnName operator ++(ColumnName cn) {
        return cn + 1;
    }
    private static uint GetNumberInternal(string name) {
        uint[] chars = new uint[3];
        chars.Initialize();
        int j = 2;
        for (int i = name.Length - 1; i >= 0; i--) {
            chars[j--] = (uint)(name[i] - 'A' + 1);
        }
        return chars[0] * 26 * 26 + chars[1] * 26 + chars[2];
    }
    private static string GetNameInternal(uint columnNumber) {
        // To store result (Excel column name)
        Stack<char> chars1 = new Stack<char>();
        while (columnNumber > 0) {
            // Find remainder
            uint rem = columnNumber % 26;
            // If remainder is 0, then a
            // 'Z' must be there in output
            if (rem == 0) {
                chars1.Push('Z');
                columnNumber = (columnNumber / 26) - 1;
            }
            // If remainder is non-zero
            else {
                chars1.Push((char)((rem - 1) + 'A'));
                columnNumber = columnNumber / 26;
            }
        }
        char[] chars2 = new char[chars1.Count];
        int i = 0;
        while (chars1.Count > 0)
            chars2[i++] = chars1.Pop();
        return new string(chars2);
    }

}
}
