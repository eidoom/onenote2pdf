/* Program : WSPBuilder
 * Created by: Carsten Keutmann
 * Date : 2007
 *  
 * The WSPBuilder comes under GNU GENERAL PUBLIC LICENSE (GPL).
 */
using System;
using System.Collections.Generic;
using System.Text;

namespace OneNote2PDF.Library
{
    public class ExceptionHandler
    {
        public static void Throw(string property, string value, string validValues)
        {
            throw new ApplicationException("Invalid value '" + value + "' for -" + property + ". Valid values are "+ validValues +".");
        }
    }
}
