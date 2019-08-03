using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Serialization;

namespace NsqlToExcel
{
    /// <summary>
    /// Common Used Extention Methods
    /// </summary>
    public static class ExtensionMethods
    {
        /// <summary>
        /// Count the number of words in a string.
        /// </summary>
        /// <param name="input">The input string.</param>
        /// <returns></returns>
        public static int WordCount(this String input)
        {
            return input.Split(new[] { ' ', '.', '?' },
                               StringSplitOptions.RemoveEmptyEntries).Length;
        }

        public static bool IsArrayOf<T>(this Type type)
        {
            return type == typeof(T[]);
        }

        public static Double ToRoundedDouble(this Decimal toRoundDecimal, Int32 places)
        {
            return Math.Round(Decimal.ToDouble(toRoundDecimal), places, MidpointRounding.AwayFromZero);
        }

        public static Double? ToRoundedDouble(this Double? toRoundDouble, Int32 places)
        {
            if (toRoundDouble == null) return null;
            return Math.Round(toRoundDouble.Value, places, MidpointRounding.AwayFromZero);
        }
        /// <summary>
        /// Lefts the specified s.
        /// </summary>
        /// <param name="s">The s.</param>
        /// <param name="length">The length.</param>
        /// <returns>String.</returns>
        public static String Left(this String s, Int32 length)
        {
            length = Math.Max(length, 0);
            return s.Length > length ? s.Substring(0, length) : s;
        }

        public static string CleanFileName(this String input)
        {
            var fileName = String.Join("", input.Split(Path.GetInvalidFileNameChars()));
            return new Regex(@"\.(?!(\w{3,4}$))").Replace(fileName, "");
        }

        public static String HexStringToString(this String hexString)
        {
            if (hexString == null || (hexString.Length & 1) == 1) return String.Empty;

            var sb = new StringBuilder();
            for (var i = 0; i < hexString.Length; i += 2)
            {
                var hexChar = hexString.Substring(i, 2);
                sb.Append((char)Convert.ToByte(hexChar, 16));
            }
            return sb.ToString();
        }


        public static String[] Lines(this String source)
        {
            return source.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
        }

        public static String Truncate(this String value, Int32 maxLength)
        {
            if (String.IsNullOrEmpty(value)) return value;
            return value.Length <= maxLength ? value : value.Substring(0, maxLength);
        }

        /// <summary>
        /// Get substring of specified number of characters on the right.
        /// </summary>
        public static String Right(this String value, int length)
        {
            return value.Substring(value.Length - length);
        }

        public static String Reverse(this String s)
        {
            var charArray = s.ToCharArray();
            Array.Reverse(charArray);
            return new string(charArray);
        }

        public static String ReplaceFirst(this String text, String search, String replace)
        {
            var pos = text.IndexOf(search, StringComparison.Ordinal);
            if (pos < 0)
            {
                return text;
            }
            return text.Substring(0, pos) + replace + text.Substring(pos + search.Length);
        }

        /// <summary>
        /// Determines whether [contains] [the specified source].
        /// </summary>
        /// <param name="source">The source.</param>
        /// <param name="toCheck">To check.</param>
        /// <param name="comp">The comp.</param>
        /// <returns>
        ///   <c>true</c> if [contains] [the specified source]; otherwise, <c>false</c>.
        /// </returns>
        public static bool Contains(this string source, string toCheck, StringComparison comp)
        {
            if (string.IsNullOrEmpty(toCheck) || string.IsNullOrEmpty(source))
                return false;

            return source.IndexOf(toCheck, comp) >= 0;
        }

        /// <summary>
        /// Converts a String to an Int32?
        /// </summary>
        /// <param name="input">The input string.</param>
        /// <returns></returns>
        public static Int32? ToInt(this string input)
        {
            int val;
            if (int.TryParse(input, out val))
                return val;
            return null;
        }

        /// <summary>
        /// Converts a string to a DateTime?
        /// </summary>
        /// <param name="input">The input string.</param>
        /// <returns></returns>
        public static DateTime? ToDate(this string input)
        {
            DateTime val;
            if (DateTime.TryParse(input, out val))
                return val;
            return null;
        }

        /// <summary>
        /// Converts a string to a Decimal?.
        /// </summary>
        /// <param name="input">The input string.</param>
        /// <returns></returns>
        public static Decimal? ToDecimal(this string input)
        {
            decimal val;
            if (decimal.TryParse(input, out val))
                return val;
            return null;
        }

        /// <summary>
        /// Converts a string to a Single?.
        /// </summary>
        /// <param name="input">The input string.</param>
        /// <returns></returns>
        public static Single? ToSingle(this string input)
        {
            Single val;
            if (Single.TryParse(input, out val))
                return val;
            return null;
        }

        /// <summary>
        /// Converts a string to a Double?.
        /// </summary>
        /// <param name="input">The input string.</param>
        /// <returns></returns>
        public static Double? ToDouble(this string input)
        {
            Double val;
            if (Double.TryParse(input, out val))
                return val;
            return null;
        }

        /// <summary>
        /// Almosts the equals.
        /// </summary>
        /// <param name="double1">The double1.</param>
        /// <param name="double2">The double2.</param>
        /// <param name="precision">The precision.</param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool AlmostEquals(this Double double1, Double double2, Double precision)
        {
            return (Math.Abs(double1 - double2) <= precision);
        }


        /// <summary>
        /// Almosts the equals.
        /// </summary>
        /// <param name="single1">The single1.</param>
        /// <param name="single2">The single2.</param>
        /// <param name="precision">The precision.</param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool AlmostEquals(this Single single1, Single single2, Single precision)
        {
            return (Math.Abs(single1 - single2) <= precision);
        }

        /// <summary>
        /// Almosts the zero.
        /// </summary>
        /// <param name="double1">The double1.</param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool AlmostZero(this Double double1)
        {
            return (Math.Abs(double1 - 0.0) <= .00000001);
        }

        /// <summary>
        /// Almosts the zero.
        /// </summary>
        /// <param name="single1">The single1.</param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool AlmostZero(this Single single1)
        {
            return Convert.ToDouble(single1).AlmostZero();
        }

        public static bool AlmostZero(this Decimal? decimal1)
        {
            return decimal1 == null || Convert.ToDouble(decimal1).AlmostZero();
        }

        public static bool AlmostZero(this Decimal decimal1)
        {
            return Convert.ToDouble(decimal1).AlmostZero();
        }

        /// <summary>
        /// Determines whether the compareTo string [is equal to] [the specified input string].
        /// </summary>
        /// <param name="input">The input string.</param>
        /// <param name="compareTo">The string to compare to.</param>
        /// <param name="ignoreCase">if set to <c>true</c> [ignore case].</param>
        /// <returns>
        ///   <c>true</c> if [is equal to] [the specified string]; otherwise, <c>false</c>.
        /// </returns>
        public static Boolean IsEqualTo(this String input, String compareTo, Boolean ignoreCase)
        {
            return string.Compare(input, compareTo, ignoreCase) == 0;
        }

        /// <summary>
        /// Determines whether the specified input string is empty.
        /// </summary>
        /// <param name="input">The input.</param>
        /// <returns>
        ///   <c>true</c> if the specified input is empty; otherwise, <c>false</c>.
        /// </returns>
        public static Boolean IsEmpty(this String input)
        {
            return (input != null && input.Length == 0);
        }

        /// <summary>
        /// Determines whether the specified input is null.
        /// </summary>
        /// <param name="input">The input string.</param>
        /// <returns>
        ///   <c>true</c> if the specified input string is null; otherwise, <c>false</c>.
        /// </returns>
        public static Boolean IsNull(this String input)
        {
            return (input == null);
        }

        /// <summary>
        /// Determines whether the input string [is null or empty].
        /// </summary>
        /// <param name="input">The input.</param>
        /// <returns>
        ///   <c>true</c> if [is null or empty] [the specified input]; otherwise, <c>false</c>.
        /// </returns>
        public static Boolean IsNullOrEmpty(this String input)
        {
            return String.IsNullOrEmpty(input);
        }

        /// <summary>
        /// Determines whether the input string [is not null or empty].
        /// </summary>
        /// <param name="input">The input string.</param>
        /// <returns>
        ///   <c>true</c> if [is not null or empty] [the specified input string]; otherwise, <c>false</c>.
        /// </returns>
        public static Boolean IsNotNullOrEmpty(this String input)
        {
            return !String.IsNullOrEmpty(input);
        }

        /// <summary>
        /// Gets a CLS complaint string based on the input string.
        /// </summary>
        /// <param name="input">The input string.</param>
        /// <returns></returns>
        public static String GetClsString(this String input)
        {
            return Regex.Replace(input.Trim(), @"[\W]", @"_");
        }

        /// <summary>
        /// Gets the name of the month.
        /// </summary>
        /// <param name="monthNumber">The month number.</param>
        /// <returns></returns>
        public static String GetMonthName(Int32 monthNumber)
        {
            return GetMonthName(monthNumber, false);
        }

        /// <summary>
        /// Gets the name of the month.
        /// </summary>
        /// <param name="monthNumber">The month number.</param>
        /// <param name="abbreviateMonth">if set to <c>true</c> [abbreviate month name].</param>
        /// <returns></returns>
        public static String GetMonthName(Int32 monthNumber, Boolean abbreviateMonth)
        {
            if (monthNumber < 1 || monthNumber > 12)
                throw new ArgumentOutOfRangeException("monthNumber");
            var date = new DateTime(1, monthNumber, 1);
            return abbreviateMonth ? date.ToString("MMM") : date.ToString("MMMM");
        }

        /// <summary>
        /// Get the Rows in a list for a the table.
        /// </summary>
        /// <param name="table">The table.</param>
        /// <returns></returns>
        public static List<DataRow> RowList(this DataTable table)
        {
            return table.Rows.Cast<DataRow>().ToList();
        }

        /// <summary>
        /// Gets the field item frp a row.
        /// </summary>
        /// <param name="row">The row.</param>
        /// <param name="field">The field.</param>
        /// <returns></returns>
        public static object GetItem(this DataRow row, String field)
        {
            return !row.Table.Columns.Contains(field) ? null : row[field];
        }

        /// <summary>
        /// Filters the rows to a list.
        /// </summary>
        /// <param name="table">The table.</param>
        /// <param name="ids">The ids.</param>
        /// <param name="fieldName">Name of the field.</param>
        /// <returns></returns>
        public static List<DataRow> FilterRowsToList(this DataTable table, List<Int32> ids, String fieldName)
        {
            Func<DataRow, bool> filter = row => ids.Contains((Int32)row.GetItem(fieldName));
            return table.RowList().Where(filter).ToList();
        }

        /// <summary>
        /// Filters the rows to a list.
        /// </summary>
        /// <param name="table">The table.</param>
        /// <param name="ids">The ids.</param>
        /// <param name="fieldName">Name of the field.</param>
        /// <returns></returns>
        public static List<DataRow> FilterRowsToList(this DataTable table, List<String> ids, String fieldName)
        {
            Func<DataRow, bool> filter = row => ids.Contains((String)row.GetItem(fieldName));
            return table.RowList().Where(filter).ToList();
        }

        /// <summary>
        /// Filters the rows to a data table.
        /// </summary>
        /// <param name="table">The table.</param>
        /// <param name="ids">The ids.</param>
        /// <param name="fieldName">Name of the field.</param>
        /// <returns></returns>
        public static DataTable FilterRowsToDataTable(this DataTable table, List<Int32> ids, String fieldName)
        {
            DataTable filteredTable = table.Clone();
            List<DataRow> matchingRows = FilterRowsToList(table, ids, fieldName);

            foreach (DataRow filteredRow in matchingRows)
            {
                filteredTable.ImportRow(filteredRow);
            }
            return filteredTable;
        }

        /// <summary>
        /// Filters the rows to a data table.
        /// </summary>
        /// <param name="table">The table.</param>
        /// <param name="ids">The ids.</param>
        /// <param name="fieldName">Name of the field.</param>
        /// <returns></returns>
        public static DataTable FilterRowsToDataTable(this DataTable table, List<String> ids, String fieldName)
        {
            DataTable filteredTable = table.Clone();
            List<DataRow> matchingRows = FilterRowsToList(table, ids, fieldName);

            foreach (DataRow filteredRow in matchingRows)
            {
                filteredTable.ImportRow(filteredRow);
            }
            return filteredTable;
        }

        /// <summary>
        /// Selects the into table the rows base on the filter.
        /// </summary>
        /// <param name="table">The table.</param>
        /// <param name="filter">The filter.</param>
        /// <returns></returns>
        public static DataTable SelectIntoTable(this DataTable table, String filter)
        {
            DataTable selectResults = table.Clone();
            DataRow[] rows = table.Select(filter);
            foreach (DataRow row in rows)
            {
                selectResults.ImportRow(row);
            }
            return selectResults;
        }

        /// <summary>
        /// Deletes rows from the table base on the filter.
        /// </summary>
        /// <param name="table">The table.</param>
        /// <param name="filter">The filter.</param>
        /// <returns></returns>
        public static DataTable Delete(this DataTable table, string filter)
        {
            table.Select(filter).Delete();
            return table;
        }

        /// <summary>
        /// Deletes the specified rows.
        /// </summary>
        /// <param name="rows">The rows.</param>
        public static void Delete(this IEnumerable<DataRow> rows)
        {
            foreach (DataRow row in rows)
                row.Delete();
        }

        /// <summary>
        /// Gets the or add value.
        /// </summary>
        /// <typeparam name="TKey">The type of the key.</typeparam>
        /// <typeparam name="TValue">The type of the value.</typeparam>
        /// <param name="dict">The dict.</param>
        /// <param name="key">The key.</param>
        /// <returns></returns>
        public static TValue GetOrAddValue<TKey, TValue>(
            this IDictionary<TKey, TValue> dict, TKey key)
                where TValue : new()
        {
            TValue value;
            if (dict.TryGetValue(key, out value))
                return value;
            value = new TValue();
            dict.Add(key, value);
            return value;
        }

        /// <summary>
        /// Gets the or add value.
        /// </summary>
        /// <typeparam name="TKey">The type of the key.</typeparam>
        /// <typeparam name="TValue">The type of the value.</typeparam>
        /// <param name="dict">The dict.</param>
        /// <param name="key">The key.</param>
        /// <param name="generator">The generator.</param>
        /// <returns></returns>
        public static TValue GetOrAddValue<TKey, TValue>(
            this IDictionary<TKey, TValue> dict, TKey key, Func<TValue> generator)
        {
            TValue value;
            if (dict.TryGetValue(key, out value))
                return value;
            value = generator();
            dict.Add(key, value);
            return value;
        }

        /// <summary>
        /// Determines whether the specified enumerable has items.
        /// </summary>
        /// <param name="enumerable">The enumerable.</param>
        /// <returns>
        ///   <c>true</c> if the specified enumerable has items; otherwise, <c>false</c>.
        /// </returns>
        public static Boolean HasItems(this IEnumerable enumerable)
        {
            if (enumerable == null) return false;
            try
            {
                var enumerator = enumerable.GetEnumerator();
                if (enumerator.MoveNext()) return true;
            }
            catch
            {
                return false;
            }
            return false;
        }

        /// <summary>
        /// Determines whether the specified enumerable is empty.
        /// </summary>
        /// <param name="enumerable">The enumerable.</param>
        /// <returns>
        ///   <c>true</c> if the specified enumerable is empty; otherwise, <c>false</c>.
        /// </returns>
        public static Boolean IsEmpty(this IEnumerable enumerable)
        {
            return !enumerable.HasItems();
        }

        public static String RemoveSpecialCharacters(this String str)
        {
            var sb = new StringBuilder();
            foreach (char c in str)
            {
                if ((c >= '0' && c <= '9') || (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || c == '.' || c == '_')
                {
                    sb.Append(c);
                }
            }
            return sb.ToString();
        }

        /// <summary>
        /// Removes the invalid chars.
        /// </summary>
        /// <param name="fileName">Name of the file.</param>
        /// <returns></returns>
        public static String RemoveInvalidChars(this String fileName)
        {
            var regex = String.Format("[{0}]", Regex.Escape(new String(Path.GetInvalidFileNameChars())));
            var removeInvalidChars = new Regex(regex,
                                               RegexOptions.Singleline | RegexOptions.Compiled |
                                               RegexOptions.CultureInvariant);
            return removeInvalidChars.Replace(fileName, "");
        }

        /// <summary>
        /// Properties the set on a given object.
        /// </summary>
        /// <param name="p">The object to set the property on.</param>
        /// <param name="propName">Name of the prop.</param>
        /// <param name="value">The value.</param>
        public static void PropertySet(this Object p, String propName, Object value)
        {
            var t = p.GetType();
            var info = t.GetProperty(propName);
            if (info == null) return;
            if (!info.CanWrite) return;
            info.SetValue(p, value, null);
        }

        /// <summary>
        /// Determines whether the specified item is between.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="item">The item.</param>
        /// <param name="start">The start.</param>
        /// <param name="end">The end.</param>
        /// <returns>
        ///   <c>true</c> if the specified item is between; otherwise, <c>false</c>.
        /// </returns>
        public static bool IsBetween<T>(this T item, T start, T end)
        {
            return Comparer<T>.Default.Compare(item, start) >= 0
                && Comparer<T>.Default.Compare(item, end) <= 0;
        }

        /// <summary>
        /// Objects to XML. UTF-8 -- use this one!
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="item">The item.</param>
        /// <returns>System.String.</returns>
        public static String ObjectToXml<T>(this T item)
        {
            var serializer = new XmlSerializer(typeof (T));
           var ns = new XmlSerializerNamespaces();
           ns.Add("", "");
           var settings = new XmlWriterSettings
           {
               Encoding = new UTF8Encoding(true),
               Indent = true,
               OmitXmlDeclaration = false,
               NewLineHandling = NewLineHandling.None
           };
            using (var stream = new MemoryStream())
            {
                using (var xmlWriter = XmlWriter.Create(stream,settings))
                {
                    serializer.Serialize(xmlWriter, item, ns);
                }
                return Encoding.UTF8.GetString(stream.ToArray()).RemoveBom();
            }
        }

        public static String RemoveBom(this String p)
        {
            var bomMarkUtf8 = Encoding.UTF8.GetString(Encoding.UTF8.GetPreamble());
            if (p.StartsWith(bomMarkUtf8)) p = p.Remove(0, bomMarkUtf8.Length);
            return p.Replace("\0", "");
        }
    }
}