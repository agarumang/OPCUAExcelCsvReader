using System;
using System.Text.RegularExpressions;

namespace ConsoleApp1
{
    public static class EncodingHelper
    {
        /// <summary>
        /// Fixes encoding issues in strings, particularly for special characters like cm³ and °C
        /// </summary>
        public static string FixEncoding(string value)
        {
            if (string.IsNullOrEmpty(value))
                return value;
            
            // Fix encoding issues: Replace cm? with cm³, ?C with °C
            // Handle various possible corrupted encodings of cm³
            value = value.Replace("cm?", "cm³");
            value = value.Replace("cm³", "cm³"); // Ensure proper superscript 3
            value = value.Replace("cm3", "cm³");
            value = value.Replace("cm^3", "cm³");
            
            // Fix temperature symbol - handle various corrupted forms
            // Replace question mark before C with degree symbol
            value = value.Replace("?C", "°C");
            value = value.Replace("? C", "°C");
            value = value.Replace("øC", "°C");
            value = value.Replace("ø C", "°C");
            value = value.Replace("ø", "°");
            value = value.Replace("°C", "°C"); // Ensure proper degree symbol
            
            // Normalize whitespace around degree symbol
            value = Regex.Replace(value, @"\s*°\s*C", "°C");
            
            return value;
        }
    }
}

