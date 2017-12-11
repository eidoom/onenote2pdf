/* Program : WSPBuilder
 * Created by: Carsten Keutmann
 * Date : 2007
 *  
 * The WSPBuilder comes under GNU GENERAL PUBLIC LICENSE (GPL).
 */
using System;
using System.Collections.Generic;
using System.Text;
using System.Diagnostics;

namespace OneNote2PDF.Library
{
    /// <summary>
    /// Log handles all output, it uses the Trace class and 
    /// therefore its possible to define more listeners in app.config
    /// </summary>
    public class Log
    {
        #region Properties 

        // Initialize and config the TraceLevelSwitch.
        private TraceSwitch _switch = null;
        protected TraceSwitch Switch
        {
            get { return _switch; }
            set { _switch = value; }
        }

        #endregion

        #region Static Methods

        public static void Error(string message)
        {
            Trace.WriteLineIf(Log.Instance.Switch.TraceError, message);
        }

        public static void Warning(string message)
        {
            Trace.WriteLineIf(Log.Instance.Switch.TraceWarning, message);
        }

        public static void Information(string message)
        {
            Trace.WriteLineIf(Log.Instance.Switch.TraceInfo, message);
        }

        public static void Verbose(string message)
        {
            Trace.WriteLineIf(Log.Instance.Switch.TraceVerbose, message);
        }

        #endregion

        #region Singleton

        protected static readonly Log instance = new Log();

        // Explicit static constructor to tell C# compiler
        // not to mark type as beforefieldinit
        static Log()
        {
        }

        Log()
        {
            Trace.Listeners.Add(new TextWriterTraceListener(Console.Out));
            
            _switch = new TraceSwitch("TraceLevelSwitch", "Trace Level for Entire Application", Config.Current.TraceLevel.ToString());
        }

        protected static Log Instance
        {
            get
            {
                return instance;
            }
        }
        #endregion
    }
}
