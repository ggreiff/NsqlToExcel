using System;
using System.CodeDom;
using System.CodeDom.Compiler;
using System.IO;
using System.Net;
using System.Web.Services.Description;
using ClosedXML.Excel;
using CommandLine;
using NLog;

namespace NsqlToExcel
{
    /// <summary>
    /// Class Program.
    /// </summary>
    public class Program
    {
        public static Logger Nlogger = LogManager.GetCurrentClassLogger();

        /// <summary>
        /// Defines the entry point of the application.
        /// </summary>
        /// <param name="args">The arguments.</param>
        public static void Main(String[] args)
        {
            Nlogger.Info("Start NsqlQueryCall");
            Nlogger.Trace("With args: {0}", String.Join(", ", args));

            var options = new Options();
            if (!Parser.Default.ParseArguments(args, options))
            {
                Nlogger.Fatal("Invalid arguments or missing required options.");
                Environment.Exit(-1);
            }

            if (options.XlsxFileName.IsNullOrEmpty()) options.XlsxFileName = $"{options.NsqlQueryCode}.xlsx";
            try
            {
                if (File.Exists(options.XlsxFileName)) File.Delete(options.XlsxFileName);
            }
            catch (Exception ex)
            {
                Nlogger.Error("Unable to delete {0}", options.XlsxFileName);
                Nlogger.Error(ex.Message);
                Nlogger.Info("Stop MainController");
                Environment.Exit(-1);
            }

            var p = new Program();
            var retVal = p.NsqlQueryCall(options);

            if (!retVal)
            {
                Nlogger.Error("NsqlQueryCall failed.");
            }
            Nlogger.Info("Stop MainController");

            if (retVal) Environment.Exit(0);
            Environment.Exit(-1);
        }

        /// <summary>
        /// NSQLs the query call.
        /// </summary>
        /// <param name="options">The options.</param>
        /// <returns>Boolean.</returns>
        private Boolean NsqlQueryCall(Options options)
        {
            var wsdl = $"https://{options.CppmHost}/niku/wsdl/Query/{options.NsqlQueryCode}?wsdl";
            Nlogger.Info("Querying {0}", wsdl);

            var webRequest = WebRequest.Create(wsdl);
            var webResponse = webRequest.GetResponse();
            ServiceDescription description = null;
            using (var stream = webResponse.GetResponseStream())
            {
                if (stream != null) description = ServiceDescription.Read(stream);
            }

            if (description == null)
            {
                Nlogger.Error("Unable to define a description");
                return false;
            }

            var importer = new ServiceDescriptionImporter
            {
                ProtocolName = "Soap",
                Style = ServiceDescriptionImportStyle.Client,
                CodeGenerationOptions = System.Xml.Serialization.CodeGenerationOptions.GenerateProperties
            };
            importer.AddServiceDescription(description, null, null);

            var codeNamespace = new CodeNamespace();
            var codeCompileUnit = new CodeCompileUnit();
            codeCompileUnit.Namespaces.Add(codeNamespace);
            var warning = importer.Import(codeNamespace, codeCompileUnit);
            if (warning != 0)
            {
                Nlogger.Warn(warning);
            }

            var codeDomProvider = CodeDomProvider.CreateProvider("C#");
            var assemblyReferences = new[] { "System.dll", "System.Web.Services.dll", "System.Web.dll", "System.Xml.dll", "System.Data.dll" };

            var compilerParameters = new CompilerParameters(assemblyReferences) { GenerateInMemory = true };
            var compileAssemblyFromDom = codeDomProvider.CompileAssemblyFromDom(compilerParameters, codeCompileUnit);
            if (compileAssemblyFromDom.Errors.Count > 0)
            {
                Nlogger.Error("compileAssemblyFromDom.Errors.Count = {0}", compileAssemblyFromDom.Errors.Count);
                foreach (CompilerError compilerError in compileAssemblyFromDom.Errors)
                {
                    Nlogger.Error(compilerError.ErrorText);
                }
                return false;
            }

            Object queryService = null;
            Object query = null;
            Object auth = null;
            //Object filter = null;
            //Object slice = null;
            //Object sort = null;

            var queryTypes = compileAssemblyFromDom.CompiledAssembly.GetTypes();
            foreach (var queryType in queryTypes)
            {
                if (queryType.FullName.IsEqualTo($"{options.NsqlQueryCode}QueryService", true)) queryService = compileAssemblyFromDom.CompiledAssembly.CreateInstance(queryType.ToString());
                //if (queryType.FullName.IsEqualTo($"{options.NsqlQueryCode}Filter", true)) filter = compileAssemblyFromDom.CompiledAssembly.CreateInstance(queryType.ToString());
                //if (queryType.FullName.IsEqualTo($"{options.NsqlQueryCode}Slice", true)) slice = compileAssemblyFromDom.CompiledAssembly.CreateInstance(queryType.ToString());
                //if (queryType.FullName.IsEqualTo($"{options.NsqlQueryCode}Sort", true)) sort = compileAssemblyFromDom.CompiledAssembly.CreateInstance(queryType.ToString());

                if (queryType.FullName.IsEqualTo("Auth", true))
                {
                    auth = compileAssemblyFromDom.CompiledAssembly.CreateInstance(queryType.ToString());
                    auth.PropertySet("Username", options.CppmUser);
                    auth.PropertySet("Password", options.CppmPassword);
                }

                if (queryType.FullName.IsEqualTo($"{options.NsqlQueryCode}Query", true))
                {
                    query = compileAssemblyFromDom.CompiledAssembly.CreateInstance(queryType.ToString());

                }
            }

            if (queryService == null)
            {
                Nlogger.Error("Unable to define a queryService");
                return false;
            }

            queryService.PropertySet("AuthValue", auth);
            query.PropertySet("Code", options.NsqlQueryCode);
            //query.PropertySet("Filter", filter);
            //query.PropertySet("Slice", slice);
            //query.PropertySet("Sort", sort);
            query.PropertySet("FilterExpression", options.FilterExpression);


            var args = new[] { query };
            var methodInfo = queryService.GetType().GetMethod("Query");
            Nlogger.Info("Invoking {0}", methodInfo.Name);
            var returnValue = methodInfo.Invoke(queryService, args);
            if (returnValue == null)
            {
                Nlogger.Error("Unable to define a Invoke returnValue");
                return false;
            }

            var propertyInfo = returnValue.GetType().GetProperty("Records");
            if (propertyInfo == null)
            {
                Nlogger.Error("Unable to define a Records propertyInfo");
                return false;
            }

            var records = propertyInfo.GetValue(returnValue, null);
            if (!(records != null && records.GetType().IsArray))
            {
                Nlogger.Error("Unable to define a Records IsArray");
                return false;
            }

            var recordsArray = records as Array;
            if (recordsArray == null || recordsArray.Length == 0)
            {
                Nlogger.Error("Unable to define a Records is empty.");
                return false;
            }
            Nlogger.Info("Processing {0} records.", recordsArray.Length);

            var rowCnt = 2;
            var colCnt = 1;
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            foreach (var record in recordsArray)
            {
                colCnt = 1;
                foreach (var prop in record.GetType().GetProperties())
                {
                    if (rowCnt == 2) ws.Cell(rowCnt - 1, colCnt).Value = prop.Name;
                    ws.Cell(rowCnt, colCnt).Value = prop.GetValue(record, null);
                    colCnt++;
                }
                rowCnt++;
            }

            ws.Columns(1, colCnt).AdjustToContents();
            foreach (var xlColumn in ws.Columns())
            {
                if (xlColumn.Width > 75) xlColumn.Width = 75;
            }
            wb.SaveAs(options.XlsxFileName);

            return true;
        }
    }
}
