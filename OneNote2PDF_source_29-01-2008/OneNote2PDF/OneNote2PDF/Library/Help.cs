/* Program : WSPBuilder
 * Created by: Carsten Keutmann
 * Date : 2007
 *  
 * The WSPBuilder comes under GNU GENERAL PUBLIC LICENSE (GPL).
 */
using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;
using System.ComponentModel;

namespace OneNote2PDF.Library
{
    public class HelpHandler
    {
        #region Const

        public const string HELP = "help";

        #endregion

        #region Members

        private Assembly _executingAssembly = null;

        private string _help = string.Empty;

        #endregion 

        #region Properties


        [DisplayName("-Help [Argument|Overview|Full] (Overview is default)")]
        [Description("Use the help to show detail description of the arguments.")]
        public string Help
        {
            get
            {
                if (string.IsNullOrEmpty(_help))
                {
                    _help = Config.Current.GetString(HELP, "OverView");
                }
                return _help;
            }
            set
            {
                _help = value;
            }
        }


        public Assembly ExecutingAssembly
        {
            get
            {
                if (_executingAssembly == null)
                {
                    _executingAssembly = Assembly.GetExecutingAssembly();
                }
                return _executingAssembly;
            }
        }

        #endregion

        #region Methods

        private Attribute GetAttribute(Type attribType)
        {
            object[] attrs = ExecutingAssembly.GetCustomAttributes(attribType, true);
            if (attrs.Length > 0)
            {
                return (Attribute)attrs[0];
            }
            return null;
        }

        private string HelpDescription(string argument)
        {
            SortedDictionary<string, string> arguments = new SortedDictionary<string, string>(StringComparer.InvariantCultureIgnoreCase);

            StringBuilder sb = new StringBuilder();
            sb.AppendLine("Arguments --------------------------------- ");

            bool validargument = false;
            bool overview = (argument.Equals("full", StringComparison.InvariantCultureIgnoreCase) ||
                            argument.Equals("overview", StringComparison.InvariantCultureIgnoreCase));

            foreach (Type type in ExecutingAssembly.GetTypes())
            {
                string category = string.Empty;
                object[] classAttribs = type.GetCustomAttributes(false);

                foreach (PropertyInfo propInfo in type.GetProperties())
                {
                    string displayName = string.Empty;
                    string description = string.Empty;

                    object[] attribs = propInfo.GetCustomAttributes(false);
                    foreach (object obj in attribs)
                    {

                        if (obj is CategoryAttribute)
                        {
                            category = ((CategoryAttribute)obj).Category;
                        }

                        if (obj is DisplayNameAttribute)
                        {
                            displayName = ((DisplayNameAttribute)obj).DisplayName;
                        }

                        if (obj is DescriptionAttribute)
                        {
                            description = ((DescriptionAttribute)obj).Description;
                        }
                    }

                    if (overview)
                    {
                        if (!string.IsNullOrEmpty(displayName))
                        {
                            string descr = string.Empty;
                            if (argument.Equals("full", StringComparison.InvariantCultureIgnoreCase))
                            {
                                descr = description;
                            }

                            arguments.Add(displayName, descr);

                            validargument = true;
                        }
                    }
                    else
                    {
                        if (displayName.StartsWith("-" + argument, StringComparison.InvariantCultureIgnoreCase))
                        {
                            // Write out only this argument
                            arguments.Add(displayName, description);

                            validargument = true;
                        }
                    }
                }
            }

            if (!validargument)
            {
                throw new ApplicationException("Invalid value '"+ Help +"' for the Help argument.");
            }

            // Write out every argument
            foreach (string key in arguments.Keys)
            {
                sb.AppendLine(key);
                if (!string.IsNullOrEmpty(arguments[key]))
                    sb.AppendLine(arguments[key]);
            }


            sb.AppendLine();
            sb.AppendLine();
            return sb.ToString();
        }

        public string Copyleft()
        {
            StringBuilder sb = new StringBuilder();

            sb.AppendLine();

            AssemblyTitleAttribute titleAttrib = (AssemblyTitleAttribute)GetAttribute(typeof(AssemblyTitleAttribute));
            sb.AppendLine(titleAttrib.Title);

            AssemblyFileVersionAttribute versionAttrib = (AssemblyFileVersionAttribute)GetAttribute(typeof(AssemblyFileVersionAttribute));
            if (versionAttrib != null)
            {
                sb.AppendLine("Version: " + versionAttrib.Version.ToString());
            }

            AssemblyCompanyAttribute companyAttrib = (AssemblyCompanyAttribute)GetAttribute(typeof(AssemblyCompanyAttribute));
            sb.AppendLine(companyAttrib.Company);
            AssemblyCopyrightAttribute copyrightAttrib = (AssemblyCopyrightAttribute)GetAttribute(typeof(AssemblyCopyrightAttribute));
            sb.AppendLine(copyrightAttrib.Copyright);

            return sb.ToString();
        }

        public string Print()
        {
            StringBuilder sb = new StringBuilder();

            sb.AppendLine(Copyleft());

            AssemblyDescriptionAttribute descriptionAttrib = (AssemblyDescriptionAttribute)GetAttribute(typeof(AssemblyDescriptionAttribute));
            sb.AppendLine(descriptionAttrib.Description);

            sb.AppendLine();

            sb.AppendLine(HelpDescription(Help));

            sb.AppendLine("Examples ----------------------------------");

            sb.AppendLine(@"OneNote2PDF -Notebook ""Shared Reference""  -CacheFolder C:\Temp\OneNote -Output D:\Temp\OneNote -ExportNotebook true");
            sb.AppendLine(@"OneNote2PDF -Notebook ""Shared Reference"" -CacheFolder C:\Temp\PDF -Output C:\Temp -ExportSection learning/programming/asp.net/MVC");
            sb.AppendLine(string.Empty);

            return sb.ToString();
        }

        #endregion

    }
}
