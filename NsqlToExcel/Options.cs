// ***********************************************************************
// Assembly         : cppmUtility
// Author           : ggreiff
// Created          : 09-01-2014
//
// Last Modified By : ggreiff
// Last Modified On : 09-01-2014
// ***********************************************************************
// <copyright file="Options.cs" company="">
//     Copyright (c) . All rights reserved.
// </copyright>
// <summary></summary>
// ***********************************************************************

using System;
using CommandLine;
using CommandLine.Text;

namespace NsqlToExcel
{
    /// <summary>
    /// Class Options.
    /// </summary>
    public class Options
    {
        /// <summary>
        /// Gets or sets the opp server.
        /// </summary>
        /// <value>The opp server.</value>
        [Option('h', "ppmHost", HelpText = "The ppm hostname.", Required = true)]
        public String CppmHost { get; set; }

        /// <summary>
        /// Gets or sets the name of the cppm user.
        /// </summary>
        /// <value>The name of the cppm user.</value>
        [Option('u', "ppmUser", HelpText = "The ppm username.", Required = true)]
        public String CppmUser { get; set; }

        /// <summary>
        /// Gets or sets the cppm password.
        /// </summary>
        /// <value>The cppm password.</value>
        [Option('p', "ppmPassword", HelpText = "The ppm user's password.", Required = true)]
        public String CppmPassword { get; set; }

        /// <summary>Gets or sets the import portfolio.</summary>
        /// <value>The import portfolio.</value>
        [Option('c', "nsqlQueryCode", HelpText = "The nsql query code to run.", Required = true)]
        public String NsqlQueryCode { get; set; }

        /// <summary>Gets or sets the import portfolio.</summary>
        /// <value>The import portfolio.</value>
        [Option('x', "xlsxFileName", HelpText = "The Xlsx FileName.  Defaults to name fo the code .xlsx")]
        public String XlsxFileName { get; set; }

        /// <summary>Gets or sets the import portfolio.</summary>
        /// <value>The import portfolio.</value>
        [Option('f', "filterExpression", HelpText = "A filter expression to use in the query. Defaults to an empty string.", DefaultValue = "")]
        public String FilterExpression { get; set; }

        /// <summary>
        /// Gets the usage.
        /// </summary>
        /// <returns>System.String.</returns>
        [HelpOption]
        public String GetUsage()
        {
            return HelpText.AutoBuild(this, current => HelpText.DefaultParsingErrorsHandler(this, current));
        }
    }
}
