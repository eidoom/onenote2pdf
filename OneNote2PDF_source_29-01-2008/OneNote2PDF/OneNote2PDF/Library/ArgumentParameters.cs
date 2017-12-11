/* Program : WSPBuilder
 *
 * This class has originally been created by: Dan H.
 * Modified by Carsten Keutmann
 * 
 * Url : http://hilres.net/(S(zsohnmqjyi5o1tislosni245))/Default.aspx?Page=ArgumentParameterParser&AspxAutoDetectCookieSupport=1
 *  
 */
using System;
using System.Collections;
using System.Collections.Generic;

namespace OneNote2PDF.Library
{
    /// <summary>
    /// The class will take a string and parse out the arguments and parameters
    /// into two collections.  The arguments must come before the parameters.
    /// Optionally it can ignore case.  If there are more then one parameters,
    /// it can ether pick the last value or append them together with a separator
    /// character.  It will try its best to parse out what’s in the string and
    /// not give any error message.
    /// </summary>
    public class ArgumentParameters : Dictionary<string, string>
    {
        //private Dictionary<string, string> parameters;
        private List<string> arguments;
        private string text;


        /// <summary>
        /// List of arguments.
        /// </summary>
        public IList<string> Arguments
        {
            get
            {
                return this.arguments;
            }
        }

        /// <summary>
        /// Text that was parsed.
        /// </summary>
        public string Text
        {
            get { return this.text; }
        }

        /// <summary>
        /// Create the ArgumentParameters class but do not parse anything yet.
        /// </summary>
        public ArgumentParameters() : base(StringComparer.CurrentCultureIgnoreCase)
        {
        }

        /// <summary>
        /// This will take a string and parse out the parameters and arguments.
        /// </summary>
        /// <param name="text">The text to parse</param>
        /// <param name="ignoreCase">True to ignore case character</param>
        /// <param name="valueSeparator">
        ///      Character to use to separate the duplicate values.
        ///      null=use last value
        ///   </param>
        public ArgumentParameters(string text, string valueSeparator) :base(StringComparer.CurrentCultureIgnoreCase)
        {
            this.Parse(text, valueSeparator);
        }

        /// <summary>
        /// Parse the commandline text.
        /// </summary>
        /// <param name="text"></param>
        /// <param name="ignoreCase"></param>
        /// <param name="valueSeparator"></param>
        public void Parse(string text, string valueSeparator)
        {
            this.text = text;
            this.arguments = new List<string>();


            int idx = 0;
            int textLength = text.Length;

            while (idx < textLength)
            {
                idx = SkipOverWhiteSpace(text, idx);

                // Do we have a parameter starter?
                if (IsParameterStarter(text[idx]))
                {// Yes, its a key.
                    idx++;
                    int leftIdx = idx;

                    // Look for a parameter / value separator.
                    while (idx < textLength)
                    {
                        if (IsParameterValueSeparator(text[idx])
                           || IsParameterStarter(text[idx])
                           || char.IsWhiteSpace(text[idx]))
                        {
                            break;
                        }
                        idx++;
                    }

                    // We got the key.
                    string key = text.Substring(leftIdx, idx - leftIdx);
                    string value;

                    idx = SkipOverWhiteSpace(text, idx);

                    // Did we bump into the next parameter?
                    if ((idx >= textLength) || IsParameterStarter(text[idx]))
                    {// Yes.
                        value = "";
                    }
                    else
                    {// No.
                        // Did we find the separator?
                        if (IsParameterValueSeparator(text[idx]))
                        {// Yes, skip over it.
                            idx++;
                            idx = SkipOverWhiteSpace(text, idx);
                        }
                        value = GetArgument(text, ref idx);
                    }

                    // Is this a duplicate key?
                    if (this.ContainsKey(key))
                    {// Yes.
                        this[key] = value;
                    }
                    else
                    {// No.
                        this.Add(key, value);
                    }
                }
                else
                {// No, its an argument.
                    this.arguments.Add(GetArgument(text, ref idx));
                }

                idx = SkipOverWhiteSpace(text, idx);
            }
        }


        /// <summary>
        /// This tests for the parameter value separator character.
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        private bool IsParameterValueSeparator(char value)
        {
            return (value == ':') || (value == '=');
        }

        /// <summary>
        /// This tests for the parameter starter character.
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        private bool IsParameterStarter(char value)
        {
            return (value == '-') || (value == '/');
        }

        /// <summary>
        /// Get an Argument.
        /// </summary>
        /// <param name="text">Text to scan over</param>
        /// <param name="idx">Starting index</param>
        /// <returns>Index past the end of the Argument or length of string</returns>
        private string GetArgument(string text, ref int idx)
        {
            int textLength = text.Length;

            // Are we at the end of the string?
            if (idx >= textLength)
            {// Yes.
                return "";
            }

            int startIdx = idx;
            char quotationChar = text[idx];

            // Is this the start of a Quotation?
            if ((quotationChar == '"') || (quotationChar == '\''))
            {// Yes.
                startIdx++;
                idx++;

                int diffCount = 0;

                while (idx < textLength)
                {
                    // Did we find a quote character?
                    if (text[idx] == quotationChar)
                    {// Yes.
                        idx++;

                        // Are there two quotes in a row?
                        if ((idx >= textLength) || (text[idx] != quotationChar))
                        {// No, we are at the end of the Quotation.
                            diffCount = -1;
                            break;
                        }
                    }
                    idx++;
                }

                string quote = quotationChar.ToString();
                return text.Substring(startIdx, idx - startIdx + diffCount).Replace(quote + quote, quote);
            }
            else
            {// No, skip until we find a separator characters.
                while (idx < textLength)
                {
                    if (Char.IsSeparator(text, idx))
                        break;
                    idx++;
                }
                return text.Substring(startIdx, idx - startIdx);
            }
        }

        /// <summary>
        /// Skip over the white space characters.
        /// </summary>
        /// <param name="text">Text to skip over</param>
        /// <param name="idx">Starting index</param>
        /// <returns>Index of non white space characher or length of string</returns>
        private int SkipOverWhiteSpace(string text, int idx)
        {
            int textLength = text.Length;
            while (idx < textLength)
            {
                if (Char.IsWhiteSpace(text, idx) == false)
                    break;
                idx++;
            }
            return idx;
        }

    }
}
