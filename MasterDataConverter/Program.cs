using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Collections;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace MasterDataConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Please Drag and Drop");
                Console.ReadKey();
                return;
            }

            if (Path.GetExtension(args[0]) == ".xlsx")
            {
                ToJson(args);
            }
            else
            {
                ToExcel(args);
            }
        }

        static void ToJson(string[] args)
        {
            Console.WriteLine("Convert To Json OK? (y/n)");
            if (Console.ReadKey().Key != ConsoleKey.Y)
            {
                return;
            }

            var dir = Path.GetDirectoryName(args[0]);
            var exportPath = dir + "/Master";
            SafeCreateDirectory(exportPath);
            XSSFWorkbook book = WorkbookFactory.Create(args[0]) as XSSFWorkbook;
            Console.WriteLine(book.NumberOfSheets);
            var defSheet = book.GetSheet("_def");
            var defs = new Dictionary<string,Dictionary<string,string>>();
            var defCells = defSheet.GetRow(0).Cells;
            var defCount = defSheet.GetRow(0).Cells.Select(_ => _.StringCellValue).Distinct().Count();
            for (var i = 0; i < defCount; i++)
            {
                var index = i * 2;
                var rowIndex = 2;
                var dic = new Dictionary<string, string>();
                while (true)
                {
                   
                    var row = defSheet.GetRow(rowIndex);
                    if (row == null)
                    {
                        break;
                    }

                    if (!row.Any())
                    {
                        break;
                    }

                    if (row.GetCell(index) == null)
                    {
                        break;
                    }

                    if (row.GetCell(index + 1) == null)
                    {
                        break;
                    }
                    //名前、数字の順
                    dic.Add(row.GetCell(index+1).StringCellValue, row.GetCell(index).StringCellValue);
                    rowIndex++;
                }
                defs.Add(defSheet.GetRow(0).GetCell(index).StringCellValue, dic);



            }
         
            for (int i = 0; i < book.NumberOfSheets; i++)
            {
                
                var rowIndex = 2;
                var sheet = book.GetSheetAt(i);
                if (sheet.SheetName == "_def")
                {
                    continue;
                }
                var defines = sheet.GetRow(0).Cells;
                var types = sheet.GetRow(1).Cells;
                var list = new List<Dictionary<string, string>>();
                
                while (true)
                {
                    var dic = new Dictionary<string, string>();
                    var row = sheet.GetRow(rowIndex);
                    if (row == null)
                    {
                        break;
                    }

                    if (!row.Any())
                    {
                        break;
                    }

                    if (row.GetCell(0) == null)
                    {
                        break;
                    }
                    for (int j = 0; j < defines.Count; j++)
                    {
                        var cell = row.GetCell(j);
                        
                        
                        if (cell != null)
                        {
                            var type = types[j].StringCellValue;
                            if (defs.ContainsKey(type))
                            {//enum
                                dic.Add(defines[j].StringCellValue, defs[type][cell.StringCellValue]);
                            }
                            else
                            {
                                switch (cell.CellType)
                                {
                                    case CellType.Unknown:
                                        dic.Add(defines[j].StringCellValue, "");
                                        break;
                                    case CellType.Numeric:
                                        dic.Add(defines[j].StringCellValue, cell.NumericCellValue.ToString());
                                        break;
                                    case CellType.String:
                                        dic.Add(defines[j].StringCellValue, cell.StringCellValue);
                                        break;
                                    case CellType.Formula:
                                        dic.Add(defines[j].StringCellValue, "");
                                        break;
                                    case CellType.Blank:
                                        dic.Add(defines[j].StringCellValue, "");
                                        break;
                                    case CellType.Boolean:
                                        dic.Add(defines[j].StringCellValue, cell.BooleanCellValue.ToString());
                                        break;
                                    case CellType.Error:
                                        break;
                                    default:
                                        throw new ArgumentOutOfRangeException();
                                }
                            }
                        }
                        else
                        {
                            dic.Add(defines[j].StringCellValue, "");
                        }
                    }
                    list.Add(dic);
                    rowIndex++;



                }
                var json = MiniJSON.Json.Serialize(list);
                File.WriteAllText(exportPath + "/" + sheet.SheetName, json);
                Console.WriteLine(exportPath + "/" + sheet.SheetName +" : " +  list.Count + " Colums");

            }
            Console.WriteLine("end");
            Console.ReadKey();
        }
        static void ToExcel(string[] args)
        {
            Console.WriteLine("Convert To Excel OK? (y/n)");
            if (Console.ReadKey().Key != ConsoleKey.Y)
            {
                return;
            }
            //Console.ReadKey();
            if (args.Length == 0)
            {
                Console.WriteLine("Please Drag and Drop");
                Console.ReadKey();
                return;
            }

            string path = args[0];
            var jsonPath = path + @"\ExportedMasterDefine";
            var masterPath = jsonPath + @"\master.json";
            var masterconfPath = jsonPath + @"\master_conf.json";
            var masterdefPath = jsonPath + @"\master_def.json";
            var masterDataPath = path + @"\Assets\Resources";
            if (
                !Directory.Exists(jsonPath) ||
                !Directory.Exists(masterDataPath) ||
                !File.Exists(masterPath) ||
                !File.Exists(masterdefPath)
            )
            {
                Console.WriteLine("Master: " + masterDataPath);
                Console.WriteLine("Json  : " + jsonPath);

                Console.WriteLine("master: " + masterPath);
                Console.WriteLine("def   : " + masterdefPath);
                Console.WriteLine("Invalid Project");
                Console.ReadKey();
                return;
            }

            Console.WriteLine("Dir  : " + path);
            Console.WriteLine("Json : " + jsonPath);
            var masterjsontext = File.ReadAllText(masterPath);
            var masterdefjsontext = File.ReadAllText(masterdefPath);
            var masterconfjsontext = File.ReadAllText(masterconfPath);
            //Console.ReadKey();

            var masterjson = MiniJSON.Json.Deserialize(masterjsontext);
            var masterdefjson = MiniJSON.Json.Deserialize(masterdefjsontext);
            var masterconfjson = MiniJSON.Json.Deserialize(masterconfjsontext);
            //Console.ReadKey();



            var bookRootPath = path + @"\ExportedMasterExcel";
            var bookPath = bookRootPath + @"\master_01.xlsx";
            var book = CreateNewBook(bookPath);


            var dataDef = masterdefjson as Dictionary<string, object>;
            var dataConf = masterconfjson as Dictionary<string, object>;
            var data = masterjson as Dictionary<string, object>;
            ICellStyle stylethin = book.CreateCellStyle();
            stylethin.BorderTop = BorderStyle.Thin;
            stylethin.BorderRight = BorderStyle.Thin;
            stylethin.BorderBottom = BorderStyle.Thin;
            stylethin.BorderTop = BorderStyle.Thin;

            foreach (var keyValuePair in data)
            {
                if (!dataConf.ContainsKey(keyValuePair.Key))
                {
                    continue;
                }
                var conf = dataConf[keyValuePair.Key] as Dictionary<string, object>;
                var sheetName = Path.GetFileName(conf["path"].ToString());
                var col = 0;
                var sheet = book.CreateSheet(sheetName);
                Console.WriteLine(keyValuePair.Key);
                var value = keyValuePair.Value as Dictionary<string, object>;
                foreach (var valuePair in value)
                {
                    WriteCell(sheet, col, 0, valuePair.Key, stylethin);
                    WriteCell(sheet, col, 1, valuePair.Value.ToString(), stylethin);
                    var key = valuePair.Value.ToString();
                    if (dataDef.ContainsKey(key))
                    {
                        //enum
                        CellRangeAddressList addressList = new CellRangeAddressList(
                            2,
                            100,
                            col,
                            col
                        );


                        var dataList = dataDef[key] as Dictionary<string, object>;
                        string[] converted = dataList.Values.ToList().ConvertAll(_ => _ as string).ToArray();
                        XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper(sheet as XSSFSheet);
                        XSSFDataValidationConstraint dvConstraint
                            = (XSSFDataValidationConstraint) dvHelper.CreateExplicitListConstraint(converted);
                        XSSFDataValidation dataValidation =
                            (XSSFDataValidation) dvHelper.CreateValidation(dvConstraint, addressList);
                        dataValidation.ShowErrorBox = true;
                        sheet.AddValidationData(dataValidation);
                    }

                    Console.WriteLine(" - " + valuePair.Key + " : " + valuePair.Value.ToString());
                    col++;
                }


                
                var p = masterDataPath + conf["path"].ToString();
                if (File.Exists(p))
                {
                    var masterdatajsontext = File.ReadAllText(p);
                    var masterdatajson = MiniJSON.Json.Deserialize(masterdatajsontext);
                    var masterList = masterdatajson as List<object>;
                    var col2 = 0;
                    var row2 = 2;
                    var defs = value.Values.ToList();
                    var intStyle = book.CreateDataFormat().GetFormat("#,##0");
                    var singleStyle = book.CreateDataFormat().GetFormat("#,##0.0");
                    foreach (object o in masterList)
                    {
                        col2 = 0;

                        var list = o as Dictionary<string, object>;
                        foreach (var valuePair in value)
                        {
                            if (list.ContainsKey(valuePair.Key))
                            {
                                var v = list[valuePair.Key].ToString();
                                var def = defs[col2].ToString();
                                if (dataDef.ContainsKey(def))
                                {
                                    //enum
                                    var dataList = dataDef[def] as Dictionary<string, object>;
                                    if (dataList.ContainsKey(v))
                                    {
                                        v = dataList[v].ToString();
                                    }
                                }

                                if (float.TryParse(v, out float arg))
                                {
                                    WriteCell(sheet, col2, row2, arg, stylethin);
                                }
                                else
                                {
                                    WriteCell(sheet, col2, row2, v, stylethin);
                                }

                                col2++;
                            }

                        }

                        row2++;
                    }
                }


            }
            var defSheet = book.CreateSheet("_def");
            var dcol = 0;
            foreach (var keyValuePair in dataDef)
            {
                WriteCell(defSheet, dcol, 0, keyValuePair.Key);
                WriteCell(defSheet, dcol + 1, 0, keyValuePair.Key);
                WriteCell(defSheet, dcol, 1, "num");
                WriteCell(defSheet, dcol + 1, 1, "name");
                var row = 2;
                var list = keyValuePair.Value as Dictionary<string, object>;
                foreach (var valuePair in list)
                {
                    WriteCell(defSheet, dcol, row, valuePair.Key);
                    WriteCell(defSheet, dcol + 1,row,valuePair.Value.ToString());
                    row++;
                }
                dcol+=2;

            }
            SafeCreateDirectory(bookRootPath);
            using (var fs = File.Create(bookPath))
            {
                book.Write(fs);
            }

            Console.ReadKey();
        }

        static IWorkbook CreateNewBook(string filePath)
        {
            IWorkbook book;
            var extension = Path.GetExtension(filePath);

            // HSSF => Microsoft Excel(xls形式)(excel 97-2003)
            // XSSF => Office Open XML Workbook形式(xlsx形式)(excel 2007以降)
            if (extension == ".xls")
            {
                book = new HSSFWorkbook();
            }
            else if (extension == ".xlsx")
            {
                book = new XSSFWorkbook();
            }
            else
            {
                throw new ApplicationException("CreateNewBook: invalid extension");
            }

            return book;
        }

        //セル設定(文字列用)
        public static ICell WriteCell(ISheet sheet, int columnIndex, int rowIndex, string value,
            ICellStyle style = null)
        {
            var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            var cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            cell.SetCellValue(value);
            if (style != null)
            {
                cell.CellStyle = style;
            }

            return cell;
        }

        //セル設定(数値用)
        public static ICell WriteCell(ISheet sheet, int columnIndex, int rowIndex, double value,
            ICellStyle style = null)
        {
            var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            var cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            cell.SetCellValue(value);
            if (style != null)
            {
                cell.CellStyle = style;
            }

            return cell;
        }

        private static DirectoryInfo SafeCreateDirectory(string path)
        {
            if (Directory.Exists(path))
            {
                return null;
            }

            return Directory.CreateDirectory(path);
        }
    }
}


/*
 * Copyright (c) 2013 Calvin Rien
 *
 * Based on the JSON parser by Patrick van Bergen
 * http://techblog.procurios.nl/k/618/news/view/14605/14863/How-do-I-write-my-own-parser-for-JSON.html
 *
 * Simplified it so that it doesn't throw exceptions
 * and can be used in Unity iPhone with maximum code stripping.
 *
 * Permission is hereby granted, free of charge, to any person obtaining
 * a copy of this software and associated documentation files (the
 * "Software"), to deal in the Software without restriction, including
 * without limitation the rights to use, copy, modify, merge, publish,
 * distribute, sublicense, and/or sell copies of the Software, and to
 * permit persons to whom the Software is furnished to do so, subject to
 * the following conditions:
 *
 * The above copyright notice and this permission notice shall be
 * included in all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
 * EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
 * MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
 * IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY
 * CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
 * TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
 * SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 */

namespace MiniJSON
{
    // Example usage:
    //
    //  using UnityEngine;
    //  using System.Collections;
    //  using System.Collections.Generic;
    //  using MiniJSON;
    //
    //  public class MiniJSONTest : MonoBehaviour {
    //      void Start () {
    //          var jsonString = "{ \"array\": [1.44,2,3], " +
    //                          "\"object\": {\"key1\":\"value1\", \"key2\":256}, " +
    //                          "\"string\": \"The quick brown fox \\\"jumps\\\" over the lazy dog \", " +
    //                          "\"unicode\": \"\\u3041 Men\u00fa sesi\u00f3n\", " +
    //                          "\"int\": 65536, " +
    //                          "\"float\": 3.1415926, " +
    //                          "\"bool\": true, " +
    //                          "\"null\": null }";
    //
    //          var dict = Json.Deserialize(jsonString) as Dictionary<string,object>;
    //
    //          Debug.Log("deserialized: " + dict.GetType());
    //          Debug.Log("dict['array'][0]: " + ((List<object>) dict["array"])[0]);
    //          Debug.Log("dict['string']: " + (string) dict["string"]);
    //          Debug.Log("dict['float']: " + (double) dict["float"]); // floats come out as doubles
    //          Debug.Log("dict['int']: " + (long) dict["int"]); // ints come out as longs
    //          Debug.Log("dict['unicode']: " + (string) dict["unicode"]);
    //
    //          var str = Json.Serialize(dict);
    //
    //          Debug.Log("serialized: " + str);
    //      }
    //  }

    /// <summary>
    /// This class encodes and decodes JSON strings.
    /// Spec. details, see http://www.json.org/
    ///
    /// JSON uses Arrays and Objects. These correspond here to the datatypes IList and IDictionary.
    /// All numbers are parsed to doubles.
    /// </summary>
    public static class Json
    {
        /// <summary>
        /// Parses the string json into a value
        /// </summary>
        /// <param name="json">A JSON string.</param>
        /// <returns>An List&lt;object&gt;, a Dictionary&lt;string, object&gt;, a double, an integer,a string, null, true, or false</returns>
        public static object Deserialize(string json)
        {
            // save the string for debug information
            if (json == null)
            {
                return null;
            }

            return Parser.Parse(json);
        }

        sealed class Parser : IDisposable
        {
            const string WORD_BREAK = "{}[],:\"";

            public static bool IsWordBreak(char c)
            {
                return Char.IsWhiteSpace(c) || WORD_BREAK.IndexOf(c) != -1;
            }

            enum TOKEN
            {
                NONE,
                CURLY_OPEN,
                CURLY_CLOSE,
                SQUARED_OPEN,
                SQUARED_CLOSE,
                COLON,
                COMMA,
                STRING,
                NUMBER,
                TRUE,
                FALSE,
                NULL
            };

            StringReader json;

            Parser(string jsonString)
            {
                json = new StringReader(jsonString);
            }

            public static object Parse(string jsonString)
            {
                using (var instance = new Parser(jsonString))
                {
                    return instance.ParseValue();
                }
            }

            public void Dispose()
            {
                json.Dispose();
                json = null;
            }

            Dictionary<string, object> ParseObject()
            {
                Dictionary<string, object> table = new Dictionary<string, object>();

                // ditch opening brace
                json.Read();

                // {
                while (true)
                {
                    switch (NextToken)
                    {
                        case TOKEN.NONE:
                            return null;
                        case TOKEN.COMMA:
                            continue;
                        case TOKEN.CURLY_CLOSE:
                            return table;
                        default:
                            // name
                            string name = ParseString();
                            if (name == null)
                            {
                                return null;
                            }

                            // :
                            if (NextToken != TOKEN.COLON)
                            {
                                return null;
                            }
                            // ditch the colon
                            json.Read();

                            // value
                            table[name] = ParseValue();
                            break;
                    }
                }
            }

            List<object> ParseArray()
            {
                List<object> array = new List<object>();

                // ditch opening bracket
                json.Read();

                // [
                var parsing = true;
                while (parsing)
                {
                    TOKEN nextToken = NextToken;

                    switch (nextToken)
                    {
                        case TOKEN.NONE:
                            return null;
                        case TOKEN.COMMA:
                            continue;
                        case TOKEN.SQUARED_CLOSE:
                            parsing = false;
                            break;
                        default:
                            object value = ParseByToken(nextToken);

                            array.Add(value);
                            break;
                    }
                }

                return array;
            }

            object ParseValue()
            {
                TOKEN nextToken = NextToken;
                return ParseByToken(nextToken);
            }

            object ParseByToken(TOKEN token)
            {
                switch (token)
                {
                    case TOKEN.STRING:
                        return ParseString();
                    case TOKEN.NUMBER:
                        return ParseNumber();
                    case TOKEN.CURLY_OPEN:
                        return ParseObject();
                    case TOKEN.SQUARED_OPEN:
                        return ParseArray();
                    case TOKEN.TRUE:
                        return true;
                    case TOKEN.FALSE:
                        return false;
                    case TOKEN.NULL:
                        return null;
                    default:
                        return null;
                }
            }

            string ParseString()
            {
                StringBuilder s = new StringBuilder();
                char c;

                // ditch opening quote
                json.Read();

                bool parsing = true;
                while (parsing)
                {

                    if (json.Peek() == -1)
                    {
                        parsing = false;
                        break;
                    }

                    c = NextChar;
                    switch (c)
                    {
                        case '"':
                            parsing = false;
                            break;
                        case '\\':
                            if (json.Peek() == -1)
                            {
                                parsing = false;
                                break;
                            }

                            c = NextChar;
                            switch (c)
                            {
                                case '"':
                                case '\\':
                                case '/':
                                    s.Append(c);
                                    break;
                                case 'b':
                                    s.Append('\b');
                                    break;
                                case 'f':
                                    s.Append('\f');
                                    break;
                                case 'n':
                                    s.Append('\n');
                                    break;
                                case 'r':
                                    s.Append('\r');
                                    break;
                                case 't':
                                    s.Append('\t');
                                    break;
                                case 'u':
                                    var hex = new char[4];

                                    for (int i = 0; i < 4; i++)
                                    {
                                        hex[i] = NextChar;
                                    }

                                    s.Append((char)Convert.ToInt32(new string(hex), 16));
                                    break;
                            }
                            break;
                        default:
                            s.Append(c);
                            break;
                    }
                }

                return s.ToString();
            }

            object ParseNumber()
            {
                string number = NextWord;

                if (number.IndexOf('.') == -1)
                {
                    long parsedInt;
                    Int64.TryParse(number, out parsedInt);
                    return parsedInt;
                }

                double parsedDouble;
                Double.TryParse(number, out parsedDouble);
                return parsedDouble;
            }

            void EatWhitespace()
            {
                while (Char.IsWhiteSpace(PeekChar))
                {
                    json.Read();

                    if (json.Peek() == -1)
                    {
                        break;
                    }
                }
            }

            char PeekChar
            {
                get
                {
                    return Convert.ToChar(json.Peek());
                }
            }

            char NextChar
            {
                get
                {
                    return Convert.ToChar(json.Read());
                }
            }

            string NextWord
            {
                get
                {
                    StringBuilder word = new StringBuilder();

                    while (!IsWordBreak(PeekChar))
                    {
                        word.Append(NextChar);

                        if (json.Peek() == -1)
                        {
                            break;
                        }
                    }

                    return word.ToString();
                }
            }

            TOKEN NextToken
            {
                get
                {
                    EatWhitespace();

                    if (json.Peek() == -1)
                    {
                        return TOKEN.NONE;
                    }

                    switch (PeekChar)
                    {
                        case '{':
                            return TOKEN.CURLY_OPEN;
                        case '}':
                            json.Read();
                            return TOKEN.CURLY_CLOSE;
                        case '[':
                            return TOKEN.SQUARED_OPEN;
                        case ']':
                            json.Read();
                            return TOKEN.SQUARED_CLOSE;
                        case ',':
                            json.Read();
                            return TOKEN.COMMA;
                        case '"':
                            return TOKEN.STRING;
                        case ':':
                            return TOKEN.COLON;
                        case '0':
                        case '1':
                        case '2':
                        case '3':
                        case '4':
                        case '5':
                        case '6':
                        case '7':
                        case '8':
                        case '9':
                        case '-':
                            return TOKEN.NUMBER;
                    }

                    switch (NextWord)
                    {
                        case "false":
                            return TOKEN.FALSE;
                        case "true":
                            return TOKEN.TRUE;
                        case "null":
                            return TOKEN.NULL;
                    }

                    return TOKEN.NONE;
                }
            }
        }

        /// <summary>
        /// Converts a IDictionary / IList object or a simple type (string, int, etc.) into a JSON string
        /// </summary>
        /// <param name="json">A Dictionary&lt;string, object&gt; / List&lt;object&gt;</param>
        /// <returns>A JSON encoded string, or null if object 'json' is not serializable</returns>
        public static string Serialize(object obj)
        {
            return Serializer.Serialize(obj);
        }

        sealed class Serializer
        {
            StringBuilder builder;

            Serializer()
            {
                builder = new StringBuilder();
            }

            public static string Serialize(object obj)
            {
                var instance = new Serializer();

                instance.SerializeValue(obj);

                return instance.builder.ToString();
            }

            void SerializeValue(object value)
            {
                IList asList;
                IDictionary asDict;
                string asStr;

                if (value == null)
                {
                    builder.Append("null");
                }
                else if ((asStr = value as string) != null)
                {
                    SerializeString(asStr);
                }
                else if (value is bool)
                {
                    builder.Append((bool)value ? "true" : "false");
                }
                else if ((asList = value as IList) != null)
                {
                    SerializeArray(asList);
                }
                else if ((asDict = value as IDictionary) != null)
                {
                    SerializeObject(asDict);
                }
                else if (value is char)
                {
                    SerializeString(new string((char)value, 1));
                }
                else
                {
                    SerializeOther(value);
                }
            }

            void SerializeObject(IDictionary obj)
            {
                bool first = true;

                builder.Append('{');

                foreach (object e in obj.Keys)
                {
                    if (!first)
                    {
                        builder.Append(',');
                    }

                    SerializeString(e.ToString());
                    builder.Append(':');

                    SerializeValue(obj[e]);

                    first = false;
                }

                builder.Append('}');
            }

            void SerializeArray(IList anArray)
            {
                builder.Append('[');

                bool first = true;

                foreach (object obj in anArray)
                {
                    if (!first)
                    {
                        builder.Append(',');
                    }

                    SerializeValue(obj);

                    first = false;
                }

                builder.Append(']');
            }

            void SerializeString(string str)
            {
                builder.Append('\"');

                char[] charArray = str.ToCharArray();
                foreach (var c in charArray)
                {
                    switch (c)
                    {
                        case '"':
                            builder.Append("\\\"");
                            break;
                        case '\\':
                            builder.Append("\\\\");
                            break;
                        case '\b':
                            builder.Append("\\b");
                            break;
                        case '\f':
                            builder.Append("\\f");
                            break;
                        case '\n':
                            builder.Append("\\n");
                            break;
                        case '\r':
                            builder.Append("\\r");
                            break;
                        case '\t':
                            builder.Append("\\t");
                            break;
                        default:
                            int codepoint = Convert.ToInt32(c);
                            if ((codepoint >= 32) && (codepoint <= 126))
                            {
                                builder.Append(c);
                            }
                            else
                            {
                                builder.Append("\\u");
                                builder.Append(codepoint.ToString("x4"));
                            }
                            break;
                    }
                }

                builder.Append('\"');
            }

            void SerializeOther(object value)
            {
                // NOTE: decimals lose precision during serialization.
                // They always have, I'm just letting you know.
                // Previously floats and doubles lost precision too.
                if (value is float)
                {
                    builder.Append(((float)value).ToString("R"));
                }
                else if (value is int
                  || value is uint
                  || value is long
                  || value is sbyte
                  || value is byte
                  || value is short
                  || value is ushort
                  || value is ulong)
                {
                    builder.Append(value);
                }
                else if (value is double
                  || value is decimal)
                {
                    builder.Append(Convert.ToDouble(value).ToString("R"));
                }
                else
                {
                    SerializeString(value.ToString());
                }
            }
        }
    }
}