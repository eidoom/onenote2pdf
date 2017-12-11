using System;
using System.ComponentModel;
using System.Configuration;
using System.Diagnostics;
using System.IO;

namespace OneNote2PDF.Library
{
    /// <summary>
    /// Serves as the configuration class for building the wsp file.
    /// </summary>
    [Category("Configuration")]
    public sealed class Config
    {
        #region Const

        private const string TRACELEVEL = "tracelevel";
        private const string NOTEBOOK = "notebook";
        private const string EXPORTSECTION = "ExportSection";
        private const string EXPORTNOTEBOOK = "ExportNotebook";
        private const string EXPORTONLY = "ExportOnly";
        private const string LISTALLNOTEBOOK = "ListAllNotebook";
        private const string EXCLUDE = "Exclude";
        private const string OUTPUT = "Output";

        private const string CACHEFOLDER = "CacheFolder";
        private const string REFRESHCACHE = "RefreshCache";

        private const string TOCLEVEL = "TOCLevel";
        private const string SHOWTOC = "ShowTOC";

        #endregion

        #region Members

        private ArgumentParameters _arguments = null;
        private bool _showHelp = false;
        private string _firstArgument = string.Empty;

        private System.Diagnostics.TraceLevel? _tracelevel = null;
        private string exportOnlySection = string.Empty;
        private string exportSection = string.Empty;

        private string _notebookName = string.Empty;
        private string[] excludeSection = null;

        #endregion

        #region Properties

        #region Config properties
        public bool ShowHelp
        {
            get { return _showHelp; }
        }


        public string FirstArgument
        {
            get { return _firstArgument; }
        }


        public ArgumentParameters Arguments
        {
            get
            {
                if (_arguments == null)
                {
                    _arguments = new ArgumentParameters();
                }
                return _arguments;
            }
        }
        #endregion

        #region Arguments

        [DisplayName("-TraceLevel [Off|Error|Warning|Info|Verbose] (Defaut value is Info)")]
        [Description("The trace level switch setting for the application. It's possible to add more Trace listeners in WSPBuilder.exe.config file.")]
        public TraceLevel TraceLevel
        {
            get
            {
                if (!_tracelevel.HasValue)
                {
                    string parameter = GetString(TRACELEVEL, "Info");
                    try
                    {
                        _tracelevel = (System.Diagnostics.TraceLevel)Enum.Parse(typeof(System.Diagnostics.TraceLevel), parameter, true);
                    }
                    catch
                    {
                        ExceptionHandler.Throw("TraceLevel", parameter, "[Off|Error|Warning|Info|Verbose]");
                    }
                }
                return _tracelevel.Value;
            }
            set { _tracelevel = value; }
        }

        [DisplayName("-Notebook [Notebook Name]")]
        [Description("Specifies the notebook name you want to export.")]
        public string NotebookName
        {
            get
            {
                _notebookName = GetString(NOTEBOOK, string.Empty);
                return _notebookName;
            }
            set { _notebookName = value; }
        }

        [DisplayName("-ExportSection [Section Name]")]
        [Description("Specifies the Section name you want to export.")]
        public string ExportSection
        {
            get
            {
                if (string.IsNullOrEmpty(exportSection))
                {
                    exportSection = GetString(EXPORTSECTION, string.Empty);
                }
                return exportSection;
            }
            set
            {
                exportSection = value;
            }
        }

        [DisplayName("-ExportOnly [Section Name]")]
        [Description("Specifies the section name you want to export. Others will not be exported")]
        public string ExportOnly
        {
            get
            {
                if (string.IsNullOrEmpty(exportOnlySection))
                {
                    exportOnlySection = GetString(EXPORTONLY, string.Empty);
                    exportSection = exportOnlySection;
                }
                return exportOnlySection;
            }
            set
            {
                exportOnlySection = value;
            }
        }

        [DisplayName("-Exclude [Section Name1, Section Name2,...]")]
        [Description("Specifies the sections to exclude from output PDF.")]
        public string[] ExcludeSections
        {
            get
            {
                if (excludeSection == null)
                {
                    string excludes = GetString(EXCLUDE, string.Empty);
                    if (string.IsNullOrEmpty(excludes))
                        return new string[0];
                    else
                    {
                        excludeSection = excludes.Split(',');
                        for (int i = 0; i < excludeSection.Length; ++i)
                        {
                            excludeSection[i] = excludeSection[i].Trim();
                        }
                        return excludeSection;
                    }
                }
                else
                    return excludeSection;
            }
            set
            {
                excludeSection = value;
            }
        }

        [DisplayName("-Output [base path name of the exported PDF]")]
        [Description("Specifies the base path name of the exported PDF.")]
        public string OutputPath
        {
            get
            {
                return GetString(OUTPUT, string.Empty);
            }
        }

        [DisplayName("-CacheFolder [base path to store all exported PDFs within one Notebook]")]
        [Description("Specifies the base path to store all exported PDFs within one Notebook.")]
        public string CacheFolder
        {
            get
            {
                return GetString(CACHEFOLDER, string.Empty);
            }
        }

        [DisplayName("-ExportNotebook [true|false] (default is false)")]
        [Description("Export entire notebook specified in -Notebook param in your current OneNote 2007.")]
        public bool ExportNotebook
        {
            get
            {
                string parameter = GetString(EXPORTNOTEBOOK, "false");
                bool result = false;
                bool.TryParse(parameter, out result);
                return result;
            }
        }

        [DisplayName("-ListAllNotebook [true|false] (default is false)")]
        [Description("List all notebooks in your current OneNote 2007.")]
        public bool ListAllNotebook
        {
            get
            {
                string parameter = GetString(LISTALLNOTEBOOK, "false");
                bool result = false;
                bool.TryParse(parameter, out result);
                return result;
            }
        }

        [DisplayName("-TOCLevel [0|1|2|...] (default is 3, zero for hiding TOC number)")]
        [Description("Maximum level at which numbering will be added to the left of TOC entry.")]
        public int TOCLevel
        {
            get
            {
                string parameter = GetString(TOCLEVEL, "3");
                int result = 3;
                int.TryParse(parameter, out result);
                return result;
            }
        }

        [DisplayName("-ShowTOC [true|false] (default is true)")]
        [Description("Specifies whether to show TOC (Table of Contents) at the begining of output PDF.")]
        public bool ShowTOC
        {
            get
            {
                string parameter = GetString(SHOWTOC, "true");
                bool result = false;
                bool.TryParse(parameter, out result);
                return result;
            }
        }

        [DisplayName("-RefreshCache [true|false] (default is false)")]
        [Description("If set, cache folder will be deleted.")]
        public bool RefreshCache
        {
            get
            {
                string parameter = GetString(REFRESHCACHE, "false");
                bool result = false;
                bool.TryParse(parameter, out result);
                return result;
            }
        }

        #endregion

        #endregion

        #region Methods
        private static Config _current = new Config();

        // Explicit static constructor to tell C# compiler
        // not to mark type as beforefieldinit
        static Config()
        {
        }

        Config()
        {
            // Parse the App.config
            // Do this parsing in the cc because it has to come before the Args Parser
            ParseAppConfig();

            // Parse the commandline.
            Arguments.Parse(Environment.CommandLine, " ");

            if (Arguments.ContainsKey("help"))
            {
                _showHelp = true;
            }
        }


        public static Config Current
        {
            get
            {
                return _current;
            }
        }

        /// <summary>
        /// Parse all the arguments from the app.config
        /// </summary>
        private void ParseAppConfig()
        {
            foreach (string key in ConfigurationManager.AppSettings.Keys)
            {
                string value = ConfigurationManager.AppSettings[key];
                if (Arguments.ContainsKey(key))
                {
                    Arguments[key] = value;
                }
                else
                {
                    Arguments.Add(key, value);
                }
            }
        }

        /// <summary>
        /// Gets a value for the specified key. A default value is defined in case that the argument has not been defined.
        /// </summary>
        /// <param name="key"></param>
        /// <param name="defaultValue"></param>
        /// <returns></returns>
        public string GetString(string key, string defaultValue)
        {
            string value = null;

            if (Arguments.ContainsKey(key))
            {
                value = Arguments[key];
            }

            if (String.IsNullOrEmpty(value))
            {
                value = defaultValue;
            }

            return value;
        }

        /// <summary>
        /// Gets a value for the specified key. A default value is defined in case that the argument has not been defined.
        /// </summary>
        /// <param name="key"></param>
        /// <param name="defaultValue"></param>
        /// <returns></returns>
        public bool GetBool(string key, bool defaultValue)
        {
            string value = this.GetString(key, defaultValue.ToString());

            bool result = defaultValue;

            bool success = bool.TryParse(value, out result);

            return result;
        }

        /// <summary>
        /// Gets a value for the specified key. A default value is defined in case that the argument has not been defined.
        /// </summary>
        /// <param name="key"></param>
        /// <param name="defaultValue"></param>
        /// <returns></returns>
        public int GetInt(string key, int defaultValue)
        {
            string value = this.GetString(key, defaultValue.ToString());

            int result = defaultValue;

            bool success = int.TryParse(value, out result);

            if (!success)
            {
                throw new ApplicationException("Invalid value '" + value + "' for key '" + key + "'. The value has to be a numeric value.");
            }

            return result;
        }

        /// <summary>
        /// Gets a value for the specified key. A default value is defined in case that the argument has not been defined.
        /// </summary>
        /// <param name="key"></param>
        /// <param name="defaultValue"></param>
        /// <returns></returns>
        public Guid GetGuid(string key, Guid defaultValue)
        {
            string value = this.GetString(key, defaultValue.ToString());

            Guid result = Guid.Empty;
            try
            {
                result = new Guid(value);
            }
            catch
            {
                throw new ApplicationException("Invalid value '" + value + "' for key '" + key + "'. The value has to be a valid GUID.");
            }
            return result;
        }

        #endregion
    }
}
