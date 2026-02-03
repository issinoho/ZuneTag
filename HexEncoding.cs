//------------------------------------------------------------------
// Zune Meta Tag Editor
// Hex Encoding Class
//
// <copyright file="HexEncoding.cs" company="The Drunken Bakery">
//     Copyright (c) 2009 The Drunken Bakery. All rights reserved.
// </copyright>
//
// Editor to update WMV meta tags for the Zune
// Encode HEX to byte array.
//
// Author: IRS
// $Revision: 1.1 $
//------------------------------------------------------------------

namespace DrunkenBakery.ZuneTag
{
    using System;
    using System.Globalization;
    using System.Linq;

    /// <summary>
    /// Summary description for HexEncoding.
    /// </summary>
    public class HexEncoding
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="HexEncoding"/> class.
        /// </summary>
        public HexEncoding()
        {
            //
            // TODO: Add constructor logic here
            //
        }

        /// <summary>
        /// Gets the byte count.
        /// </summary>
        /// <param name="hexString">The hex string.</param>
        /// <returns></returns>
        public static int GetByteCount(string hexString)
        {
            // Remove all none A-F, 0-9, characters
            var numHexChars = hexString.Count(IsHexDigit);

            // If odd number of characters, discard last character
            if (numHexChars % 2 != 0) numHexChars--;

            return numHexChars / 2; // 2 characters per byte
        }

        /// <summary>
        /// Creates a byte array from the hexadecimal string. Each two characters are combined
        /// to create one byte. First two hexadecimal characters become first byte in returned array.
        /// Non-hexadecimal characters are ignored. 
        /// </summary>
        /// <param name="hexString">string to convert to byte array</param>
        /// <param name="discarded">number of characters in string ignored</param>
        /// <returns>byte array, in the same left-to-right order as the hexString</returns>
        public static byte[] GetBytes(string hexString, out int discarded)
        {
            discarded = 0;
            var newString = "";
            // Remove all none A-F, 0-9, characters
            foreach (var t in hexString)
                if (IsHexDigit(t))
                    newString += t;
                else
                    discarded++;

            // If odd number of characters, discard last character
            if (newString.Length % 2 != 0)
            {
                discarded++;
                newString = newString.Substring(0, newString.Length - 1);
            }

            var byteLength = newString.Length / 2;
            var bytes = new byte[byteLength];
            var j = 0;
            for (var i = 0; i < bytes.Length; i++)
            {
                var hex = new string(new[] { newString[j], newString[j + 1] });
                bytes[i] = HexToByte(hex);
                j += 2;
            }

            return bytes;
        }

        /// <summary>
        /// Toes the string.
        /// </summary>
        /// <param name="bytes">The bytes.</param>
        /// <returns></returns>
        public static string ToString(byte[] bytes)
        {
            return bytes.Aggregate("", (current, t) => current + t.ToString("X2"));
        }

        /// <summary>
        /// Determines if given string is in proper hexadecimal string format
        /// </summary>
        /// <param name="hexString"></param>
        /// <returns></returns>
        public static bool InHexFormat(string hexString)
        {
            return hexString.All(IsHexDigit);
        }

        /// <summary>
        /// Returns true is c is a hexadecimal digit (A-F, a-f, 0-9)
        /// </summary>
        /// <param name="c">Character to test</param>
        /// <returns>true if hex digit, false if not</returns>
        public static bool IsHexDigit(char c)
        {
            var numA = Convert.ToInt32('A');
            var num1 = Convert.ToInt32('0');
            c = char.ToUpper(c);
            var numChar = Convert.ToInt32(c);
            if (numChar >= numA && numChar < numA + 6)
                return true;
            return numChar >= num1 && numChar < num1 + 10;
        }

        /// <summary>
        /// Converts 1 or 2 character string into equivalent byte value
        /// </summary>
        /// <param name="hex">1 or 2 character string</param>
        /// <returns>byte</returns>
        private static byte HexToByte(string hex)
        {
            if (hex.Length > 2 || hex.Length <= 0)
                throw new ArgumentException("hex must be 1 or 2 characters in length");
            var newByte = byte.Parse(hex, NumberStyles.HexNumber);
            return newByte;
        }
    }
}