// HLUTool is used to view and maintain habitat and land use GIS data.
// Copyright © 2011 Hampshire Biodiversity Information Centre
// Copyright © 2013 Thames Valley Environmental Records Centre
// 
// This file is part of HLUTool.
// 
// HLUTool is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
// 
// HLUTool is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
// 
// You should have received a copy of the GNU General Public License
// along with HLUTool.  If not, see <http://www.gnu.org/licenses/>.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using HLU.Properties;

namespace HLU.Date
{
    public class VagueDateInstance
    {
        public int StartDate;
        public int EndDate;
        public string DateType;
        public string UserEntry;

        public bool IsUnknown
        {
            get
            {
                return String.IsNullOrEmpty(this.DateType) ||
                    (this.DateType == VagueDate.ToCode(VagueDate.VagueDateTypes.Unknown)) && !this.IsBad;
            }
        }

        public bool IsBad
        {
            get
            {
                return String.IsNullOrEmpty(this.DateType) || (this.DateType == VagueDate.ToCode(VagueDate.VagueDateTypes.Unknown)) &&
                    !VagueDate.IsUnknownDate(this.UserEntry);
            }
        }

        public VagueDateInstance(int startDate, int endDate, string dateType)
        {
            StartDate = startDate;
            EndDate = endDate;
            DateType = dateType;
            UserEntry = null;
        }

        public VagueDateInstance(int startDate, int endDate, string dateType, string userEntry)
        {
            StartDate = startDate;
            EndDate = endDate;
            DateType = dateType;
            UserEntry = userEntry;
        }

        public VagueDateInstance()
        {
            StartDate = -693593;
            EndDate = -693593;
            DateType = null;
        }
    }

    class VagueDate
    {
        private static Dictionary<string, VagueDate.VagueDateTypes> _codeValueDictionary = Enum<VagueDate.VagueDateTypes>.ToCodeValueDictionary();
        private static Dictionary<VagueDate.VagueDateTypes, string> _valueCodeDictionary = Enum<VagueDate.VagueDateTypes>.ToValueCodeDictionary();
        private static Dictionary<string, string> _codeNameDictionary = Enum<VagueDate.VagueDateTypes>.ToCodeNameDictionary();
        private static Dictionary<string, string> _nameCodeDictionary = Enum<VagueDate.VagueDateTypes>.ToNameCodeDictionary();

        /// <summary>
        /// Base date for vague dates (COM base date).
        /// </summary>
        public static DateTime BaseDate = new DateTime(1899, 12, 30);

        /// <summary>
        /// Time difference in days between COM base date (30/12/1899) and .NET base date (1/1/1) 
        /// returned by ToTimeSpanDays overloads when conversion fails and date is unkown.
        /// </summary>
        //public static readonly int DateUnknown = -693593;
        public static readonly int DateUnknown = 0;

        public enum DateType { Start, End, Vague };

        public enum VagueDateTypes
        {
            [EnumCode("D")]
            StartDate,
            [EnumCode("DD")]
            StartAndEndDates,
            [EnumCode("D-")]
            StartDateRange,
            [EnumCode("-D")]
            EndDateRange,
            [EnumCode("O")]
            StartMonthAndYear,
            [EnumCode("OO")]
            StartAndEndMonthAndYear,
            [EnumCode("O-")]
            StartMonthRange,
            [EnumCode("-O")]
            EndMonthRange,
            [EnumCode("Y")]
            StartYear,
            [EnumCode("YY")]
            StartAndEndYear,
            [EnumCode("Y-")]
            StartYearRange,
            [EnumCode("-Y")]
            EndYearRange,
            [EnumCode("P")]
            StartSeason,
            [EnumCode("PP")]
            StartAndEndSeason,
            [EnumCode("P-")]
            StartSeasonRange,
            [EnumCode("-P")]
            EndSeasonRange,
            [EnumCode("U")]
            Unknown
        }

        public static string Delimiter = Settings.Default.VagueDateDelimiter;

        public static string[] SeasonNames = Settings.Default.SeasonNames.Cast<string>().ToArray();

        public static string ToCode(VagueDateTypes en)
        {
            string outString;
            _valueCodeDictionary.TryGetValue(en, out outString);
            return outString ?? null;
        }

        public static VagueDateTypes FromCode(string code)
        {
            VagueDateTypes vd;
            if (_codeValueDictionary.TryGetValue(code, out vd))
                return vd;
            else
                return VagueDateTypes.Unknown;
        }

        /// <summary>
        /// Determines the format of a vague date string and standardizes its format.
        /// </summary>
        /// <param name="userDateString">Date string to be </param>
        /// <param name="formattedDate">Formatted version of userDateString.</param>
        /// <returns>Format string that describes the format of the date passed in:
        /// "D":  start date
        /// "DD": start date + end date
        /// "D-": start date range
        /// "-D": end date range
        /// "O":  start month and year
        /// "OO": start month and year + end month and year
        /// "O-": start month range
        /// "-O": end month range
        /// "Y":  year
        /// "YY": start year + end year
        /// "Y-": start year range
        /// "-Y": end year range
        /// "P":  season
        /// "PP": start season + end season
        /// "P-": start season range
        /// "-P": end season range
        /// "U":  unknown
        /// An empty string is returned if userDateString is null or empty. Unknown if an error occurs.
        /// </returns>
        public static string GetType(string userDateString, out string formattedDate)
        {
            // Default the output to the input.
            formattedDate = userDateString;

            // Return an empty string is null or empty.
            if (String.IsNullOrEmpty(userDateString)) return String.Empty;

            string startDateString;
            string endDateString;
            bool delimiterFound;
            // Split the string into start and end parts (the end part will be null
            // if there is only one part.
            SplitDateString(userDateString, out startDateString, out endDateString, out delimiterFound);

            string formatString1 = String.Empty;

            // If no delimeter was found and the end part is null return just the formatted first part
            if (!delimiterFound && String.IsNullOrEmpty(endDateString))
            {
                if (IsUnknownDate(startDateString))
                {
                    formattedDate = VagueDateTypes.Unknown.ToString();
                    return VagueDate.ToCode(VagueDateTypes.Unknown);
                }
                else
                {
                    return ReadDateString(startDateString, String.Empty, out formattedDate);
                }
            }
            else
            {
                // Format the second part of the date
                string formatString2 = ReadDateString(endDateString, String.Empty, out endDateString);

                // Get the year of the second part
                string endYearStr = String.Empty;
                int endYear = -1;
                if ((endDateString.Length >= 4) && Int32.TryParse(endDateString.Substring(endDateString.Length - 4, 4),
                    out endYear)) endYearStr = endYear.ToString();

                // Format the first part of the date (using the year from the end part ???)
                formatString1 = ReadDateString(startDateString, endYearStr, out startDateString);

                // Clear the formatted first or second parts if they are unknown (and the other part isn't)
                if ((formatString1 == VagueDate.ToCode(VagueDateTypes.Unknown)) && (formatString2 != VagueDate.ToCode(VagueDateTypes.Unknown)))
                    formatString1 = String.Empty;
                else if ((formatString1 != VagueDate.ToCode(VagueDateTypes.Unknown)) && (formatString2 == VagueDate.ToCode(VagueDateTypes.Unknown)))
                    formatString2 = String.Empty;

                // IF both parts are unknown return an unknown date
                if ((formatString1 == VagueDate.ToCode(VagueDateTypes.Unknown)) && (formatString2 == VagueDate.ToCode(VagueDateTypes.Unknown)))
                {
                    formattedDate = VagueDateTypes.Unknown.ToString();
                    return VagueDate.ToCode(VagueDateTypes.Unknown);
                }

                // Return the formatted first and second parts combined together
                if (!String.IsNullOrEmpty(formatString1) && String.IsNullOrEmpty(formatString2))
                {
                    formattedDate = startDateString + Delimiter;
                    return formatString1 + Delimiter;
                }
                else if (String.IsNullOrEmpty(formatString1) && !String.IsNullOrEmpty(formatString2))
                {
                    formattedDate = Delimiter + endDateString;
                    return Delimiter + formatString2;
                }
                else if (startDateString != endDateString)
                {
                    formattedDate = startDateString + Delimiter + endDateString;
                    return formatString1 + formatString2;
                }
                else
                {
                    formattedDate = startDateString;
                    return formatString1;
                }
            }
        }

        /// <summary>
        /// Determines whether the specified date string is unknown (i.e. equals null or "U").
        /// </summary>
        /// <param name="dateString">The date string.</param>
        /// <returns>
        ///   <c>true</c> if the specified date string is unknown; otherwise, <c>false</c>.
        /// </returns>
        public static bool IsUnknownDate(string dateString)
        {
            string unkString = VagueDateTypes.Unknown.ToString();
            return (dateString == null) || ((dateString.Length <= unkString.Length) && unkString.Substring(0, dateString.Length)
                .Equals(dateString, StringComparison.CurrentCultureIgnoreCase));
        }

        /// <summary>
        /// Reads one portion, either start or end, of an unformatted vague date string (i.e. after splitting).
        /// </summary>
        /// <param name="dateString">Start or end portion of a vague date string.</param>
        /// <param name="endYear">A single integer could be both a year or a month, 
        /// and the start part of date ranges may not include the year, the end date part
        /// of a vague date string is read first and the year found is passed into this method.</param>
        /// <param name="formattedDate">Formatted version of userDateString.</param>
        /// <returns>Single letter format string for one portion of a vague date.</returns>
        private static string ReadDateString(string dateString, string endYear, out string formattedDate)
        {
            formattedDate = dateString;

            try
            {
                if (String.IsNullOrEmpty(dateString)) return String.Empty;

                if (dateString.ToLower() == "unknown") return VagueDateTypes.Unknown.ToString();

                DateTimeFormatInfo dtFormatInfo = CultureInfo.CurrentCulture.DateTimeFormat;

                string[] splitArray = Regex.Split(dateString, @"[\s+" + dtFormatInfo.DateSeparator + "]");

                int year = Int32.MinValue;
                int month = Int32.MinValue;
                string currYear = String.IsNullOrEmpty(endYear) ? DateTime.Now.Year.ToString() : endYear;
                string userYear;

                switch (splitArray.Length)
                {
                    case 1:
                        if (Int32.TryParse(splitArray[0], out year))
                        {
                            if (String.IsNullOrEmpty(endYear) || (year > 12) || (year < 0))
                            {
                                // year ("2007" or "7")
                                formattedDate = new DateTime(year, 1, 1).Year.ToString();
                                return VagueDate.ToCode(VagueDateTypes.StartYear); // "Y";
                            }
                            else
                            {
                                // first part of a month range ("1-3/7")
                                formattedDate = String.Format("{0} {1}", dtFormatInfo.MonthNames[year - 1], endYear);
                                return VagueDate.ToCode(VagueDateTypes.StartMonthAndYear); // "O";
                            }
                        }
                        else if (dtFormatInfo.MonthNames.Count(s => s.ToLower() == splitArray[0].ToLower()) == 1)
                        {
                            // full month ("January")
                            formattedDate = String.Format("{0} {1}",
                                dtFormatInfo.MonthNames.Single(s => s.ToLower() == splitArray[0].ToLower()), currYear);
                            return VagueDate.ToCode(VagueDateTypes.StartMonthAndYear); // "O";
                        }
                        else if (dtFormatInfo.AbbreviatedMonthNames.Count(s => s.ToLower() == splitArray[0].ToLower()) == 1)
                        {
                            // abbreviated month ("Jan")
                            formattedDate = String.Format("{0} {1}", dtFormatInfo.MonthNames[System.Array.IndexOf(
                                dtFormatInfo.AbbreviatedMonthNames, dtFormatInfo.AbbreviatedMonthNames.Single(
                                s => s.ToLower() == splitArray[0].ToLower()))], currYear);
                            return VagueDate.ToCode(VagueDateTypes.StartMonthAndYear); // "O";
                        }
                        else if (SeasonNames.Count(s => s.ToLower() == splitArray[0].ToLower()) == 1)
                        {
                            // season
                            formattedDate = String.Format("{0} {1}",
                                SeasonNames.Single(s => s.ToLower() == splitArray[0].ToLower()), currYear);
                            return VagueDate.ToCode(VagueDateTypes.StartSeason); // "P";
                        }
                        break;
                    case 2:
                        if (!Int32.TryParse(splitArray[1], out year))
                            throw new ArgumentException("Invalid date string.", "userDateString");
                        userYear = new DateTime(year, 1, 1).Year.ToString();
                        if (Int32.TryParse(splitArray[0], out month))
                        {
                            formattedDate = String.Format("{0} {1}", dtFormatInfo.MonthNames[month - 1], userYear);
                            return VagueDate.ToCode(VagueDateTypes.StartMonthAndYear); // "O";
                        }
                        else if (dtFormatInfo.MonthNames.Count(s => s.ToLower() == splitArray[0].ToLower()) == 1)
                        {
                            // full month and year ("January 2007")
                            formattedDate = String.Format("{0} {1}",
                                dtFormatInfo.MonthNames.Single(s => s.ToLower() == splitArray[0].ToLower()), userYear);
                            return VagueDate.ToCode(VagueDateTypes.StartMonthAndYear); // "O";
                        }
                        else if (dtFormatInfo.AbbreviatedMonthNames.Count(s => s.ToLower() == splitArray[0].ToLower()) == 1)
                        {
                            // abbreviated month and year("Jan 2007")
                            formattedDate = String.Format("{0} {1}", dtFormatInfo.MonthNames[System.Array.IndexOf(
                                dtFormatInfo.AbbreviatedMonthNames, dtFormatInfo.AbbreviatedMonthNames.Single(
                                s => s.ToLower() == splitArray[0].ToLower()))], userYear);
                            return VagueDate.ToCode(VagueDateTypes.StartMonthAndYear); // "O";
                        }
                        else if (SeasonNames.Count(s => s.ToLower() == splitArray[0].ToLower()) == 1)
                        {
                            // season
                            formattedDate = String.Format("{0} {1}",
                                SeasonNames.Single(s => s.ToLower() == splitArray[0].ToLower()), userYear);
                            return VagueDate.ToCode(VagueDateTypes.StartSeason); // "P";
                        }
                        break;
                }

                DateTime testDate;

                if (DateTime.TryParse(dateString, dtFormatInfo, DateTimeStyles.AllowWhiteSpaces |
                    DateTimeStyles.AssumeLocal | DateTimeStyles.NoCurrentDateDefault, out testDate) &&
                    (testDate.Year != 1)) return VagueDate.ToCode(VagueDateTypes.StartDate); // "D" // normal date
            }
            catch { }

            return String.Empty;
        }

        /// <summary>
        /// Creates a vague date string from a vague date instance containing a date type start and end dates 
        /// expressed as days elapsed since 30/12/1899.
        /// </summary>
        /// <param name="vd">a VagueDateInstance with the date to be converted.</param>
        /// <param name="outputFormat">A vague date string in the format specified in the dateType parameter.</param>
        /// <returns></returns>
        public static string FromVagueDateInstance(VagueDateInstance vd, DateType outputFormat)
        {
            return FromTimeSpanDays(vd.StartDate, vd.EndDate, vd.DateType, outputFormat);
        }

        /// <summary>
        /// Creates a vague date string of a specified format from start and end dates 
        /// expressed as days elapsed since 30/12/1899.
        /// </summary>
        /// <param name="startDateDays">Start date in days since 30/12/1899.</param>
        /// <param name="endDateDays">End date in days since 30/12/1899.</param>
        /// <param name="dateType">Two-letter date format code as generated by the GetFormat method.</param>
        /// <param name="outputFormat">One of the output formats in the HLU.Date.DateType enumeration.</param>
        /// <returns>A vague date string in the format specified in the dateType parameter.</returns>
        public static string FromTimeSpanDays(int startDateDays, int endDateDays,
            string dateType, DateType outputFormat)
        {
            try
            {
                bool useString = false;
                string startDateString = null;
                string endDateString = null;
                if (startDateDays == Int32.MinValue)
                {
                    startDateString = String.Empty;
                    endDateString = BaseDate.AddDays(endDateDays).ToShortDateString();
                    useString = true;
                }
                if (endDateDays == Int32.MinValue)
                {
                    startDateString = BaseDate.AddDays(startDateDays).ToShortDateString();
                    endDateString = String.Empty;
                    useString = true;
                }
                if (!useString)
                    return FromDate(BaseDate.AddDays(startDateDays),
                        BaseDate.AddDays(endDateDays), dateType, outputFormat);
                else
                    return FromDateString(startDateString, endDateString, dateType, outputFormat);
            }
            catch { return FromDateString(String.Empty, String.Empty, dateType, outputFormat); }
        }

        /// <summary>
        /// Creates a vague date string of a specified format from start and end dates in ticks.
        /// </summary>
        /// <param name="startDateTicks">Start date in ticks.</param>
        /// <param name="endDateTicks">End date in ticks.</param>
        /// <param name="dateType">Two-letter date format code as generated by the GetFormat method.</param>
        /// <param name="outputFormat">One of the output formats in the HLU.Date.DateType enumeration.</param>
        /// <returns>A vague date string in the format specified in the dateType parameter.</returns>
        public static string FromTimeSpanTicks(long startDateTicks, long endDateTicks,
            string dateType, DateType outputFormat)
        {
            try
            {
                DateTime startDate = new DateTime(startDateTicks);
                DateTime endDate = new DateTime(endDateTicks);
                return FromDate(startDate, endDate, dateType, outputFormat);
            }
            catch { return FromDateString(String.Empty, String.Empty, dateType, outputFormat); }
        }

        /// <summary>
        /// Creates a vague date string of a specified format from a date string.
        /// </summary>
        /// <param name="dateString">Date string.</param>
        /// <param name="dateType">Two-letter date format code as generated by the GetFormat method.</param>
        /// <param name="outputFormat">One of the output formats in the HLU.Date.DateType enumeration.</param>
        /// <returns>A vague date string in the format specified in the dateType parameter.</returns>
        public static string FromDateString(string dateString, string dateType, DateType outputFormat)
        {
            string startDateString;
            string endDateString;
            bool delimiterFound;
            SplitDateString(dateString, out startDateString, out endDateString, out delimiterFound);
            return FromDateString(startDateString, endDateString, dateType, outputFormat);
        }

        /// <summary>
        /// Creates a vague date string of a specified format from start and end date strings.
        /// </summary>
        /// <param name="startDateString"></param>
        /// <param name="endDateString"></param>
        /// <param name="dateType">Two-letter date format code as generated by the GetFormat method.</param>
        /// <param name="outputFormat">One of the output formats in the HLU.Date.DateType enumeration.</param>
        /// <returns>A vague date string in the format specified in the dateType parameter.</returns>
        public static string FromDateString(string startDateString, string endDateString,
            string dateType, DateType outputFormat)
        {
            try
            {
                DateTimeFormatInfo dtFormatInfo = CultureInfo.CurrentCulture.DateTimeFormat;

                DateTime startDate = DateTime.MinValue;
                DateTime endDate = DateTime.MinValue;

                if ((String.IsNullOrEmpty(startDateString) || !DateTime.TryParse(startDateString, dtFormatInfo,
                    DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AssumeLocal |
                    DateTimeStyles.NoCurrentDateDefault, out startDate)) && ((outputFormat == DateType.Start) ||
                    (outputFormat == DateType.Vague))) return VagueDateTypes.Unknown.ToString();

                if ((String.IsNullOrEmpty(endDateString) || !DateTime.TryParse(endDateString, dtFormatInfo,
                    DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AssumeLocal |
                    DateTimeStyles.NoCurrentDateDefault, out endDate)) && ((outputFormat == DateType.End) ||
                    (outputFormat == DateType.Vague))) return VagueDateTypes.Unknown.ToString();

                return FromDate(startDate, endDate, dateType, outputFormat);
            }
            catch { return FromDateString(String.Empty, String.Empty, dateType, outputFormat); }
        }

        /// <summary>
        /// Creates a vague date string of a specified format from start and end dates.
        /// </summary>
        /// <param name="startDate">Start date.</param>
        /// <param name="endDate">End date.</param>
        /// <param name="dateType">Two-letter date format code as generated by the GetFormat method.</param>
        /// <param name="outputFormat">One of the output formats in the HLU.Date.DateType enumeration.</param>
        /// <returns>A vague date string in the format specified in the dateType parameter.</returns>
        public static string FromDate(DateTime startDate, DateTime endDate,
            string dateType, DateType outputFormat)
        {
            if (String.IsNullOrEmpty(dateType) || (dateType == VagueDate.ToCode(VagueDateTypes.Unknown)))
                return VagueDateTypes.Unknown.ToString();

            try
            {
                DateTimeFormatInfo dtFormatInfo = CultureInfo.CurrentCulture.DateTimeFormat;

                string formatString = String.Empty;
                switch (outputFormat)
                {
                    case DateType.Start:
                        // Use the start date type to format the start date.
                        if (dateType.Length > 0)
                            formatString = dateType.Substring(0, 1);
                        switch (VagueDate.FromCode(formatString))
                        {
                            case VagueDateTypes.StartDate: // "D"
                                return startDate.ToShortDateString();
                            case VagueDateTypes.StartMonthAndYear: // "O"
                                return String.Format("{0} {1}", dtFormatInfo.MonthNames[startDate.Month - 1], startDate.Year);
                            case VagueDateTypes.StartYear: // "Y"
                                return startDate.Year.ToString();
                            case VagueDateTypes.StartSeason: // "P"
                                return String.Format("{0}{1}", SeasonString(startDate), Delimiter);
                            case VagueDateTypes.Unknown: // "U"
                                return VagueDateTypes.Unknown.ToString();
                        }
                        break;
                    case DateType.End:
                        if (dateType.Length > 1)
                            formatString = dateType.Substring(dateType.Length - 1, 1);
                        switch (VagueDate.FromCode(formatString))
                        {
                            case VagueDateTypes.StartDate: // "D"
                                return endDate.ToShortDateString();
                            case VagueDateTypes.StartMonthAndYear: // "O"
                                return String.Format("{0} {1}", dtFormatInfo.MonthNames[endDate.Month - 1], endDate.Year);
                            case VagueDateTypes.StartYear: // "Y"
                                return endDate.Year.ToString();
                            case VagueDateTypes.StartSeason: // "P"
                                return String.Format("{0}{1}", Delimiter, SeasonString(endDate));
                            case VagueDateTypes.Unknown: // "U"
                                return VagueDateTypes.Unknown.ToString();
                        }
                        break;
                    case DateType.Vague:
                        if (dateType.Length > 0)
                            formatString = dateType.Substring(0, 1);

                        StringBuilder returnString = new StringBuilder();

                        switch (VagueDate.FromCode(formatString))
                        {
                            case VagueDateTypes.StartDate: // "D"
                                returnString.Append(startDate.ToShortDateString());
                                break;
                            case VagueDateTypes.StartMonthAndYear: // "O"
                                returnString.Append(String.Format("{0} {1}",
                                    dtFormatInfo.MonthNames[startDate.Month - 1], startDate.Year));
                                break;
                            case VagueDateTypes.StartYear: // "Y"
                                returnString.Append(startDate.Year.ToString());
                                break;
                            case VagueDateTypes.StartSeason: // "P"
                                returnString.Append(SeasonString(startDate));
                                break;
                            case VagueDateTypes.Unknown: // "U"
                                break;
                        }

                        if ((endDate == startDate) || (dateType.Length < 2))
                            return returnString.ToString();

                        formatString = dateType.Substring(dateType.Length - 1, 1);
                        switch (VagueDate.FromCode(formatString))
                        {
                            case VagueDateTypes.StartDate: // "D"
                                return returnString.Append(Delimiter).Append(endDate.ToShortDateString()).ToString();
                            case VagueDateTypes.StartMonthAndYear: // "O"
                                return returnString.Append(Delimiter).Append(String.Format("{0} {1}",
                                    dtFormatInfo.MonthNames[endDate.Month - 1], endDate.Year)).ToString();
                            case VagueDateTypes.StartYear: // "Y"
                                return returnString.Append(Delimiter).Append(endDate.Year).ToString();
                            case VagueDateTypes.StartSeason: // "P"
                                return returnString.Append(Delimiter).Append(SeasonString(endDate)).ToString();
                            case VagueDateTypes.Unknown: // "U"
                                if (returnString.Length == 0)
                                    return VagueDateTypes.Unknown.ToString();
                                else
                                    return returnString.Append(Delimiter).ToString();
                        }
                        break;
                }
            }
            catch { }

            return VagueDateTypes.Unknown.ToString();
        }

        /// <summary>
        /// Converts a vague date string into a timespan in days since 30/12/1899.
        /// </summary>
        /// <param name="dateString">Vague date string.</param>
        /// <param name="dateType">Two-letter date format code as generated by the GetFormat method.</param>
        /// <param name="outputFormat">One of the output formats in the HLU.Date.DateType enumeration.</param>
        /// <returns></returns>
        public static int ToTimeSpanDays(string dateString, string dateType, DateType outputFormat)
        {
            string startDateString;
            string endDateString;
            bool delimiterFound;
            SplitDateString(dateString, out startDateString, out endDateString, out delimiterFound);

            if (outputFormat == DateType.End && endDateString == null)
                endDateString = startDateString;
            return ToTimeSpanDays(startDateString, endDateString, dateType, outputFormat);
        }

        /// <summary>
        /// Converts a vague date composed of separate start and end date strings into a timespan in days since 30/12/1899.
        /// </summary>
        /// <param name="startDateString">Start date string.</param>
        /// <param name="endDateString">End date string.</param>
        /// <param name="dateType">Two-letter date format code as generated by the GetFormat method.</param>
        /// <param name="outputFormat">One of the output formats in the HLU.Date.DateType enumeration.</param>
        /// <returns></returns>
        public static int ToTimeSpanDays(string startDateString, string endDateString,
            string dateType, DateType outputFormat)
        {
            if (String.IsNullOrEmpty(dateType) || (dateType == VagueDate.ToCode(VagueDateTypes.Unknown)))
                return DateUnknown;

            try
            {
                DateTimeFormatInfo dtFormatInfo = CultureInfo.CurrentCulture.DateTimeFormat;

                DateTime startDate;
                bool startDateOk = DateTime.TryParse(startDateString, dtFormatInfo,
                    DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AssumeLocal |
                    DateTimeStyles.NoCurrentDateDefault, out startDate);

                DateTime endDate;
                bool endDateOk = DateTime.TryParse(endDateString, dtFormatInfo,
                    DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AssumeLocal |
                    DateTimeStyles.NoCurrentDateDefault, out endDate);

                int year;

                string formatType = String.Empty;
                switch (outputFormat)
                {
                    case DateType.Start:
                        if (dateType.Length > 0)
                            formatType = dateType.Substring(0, 1);
                        switch (VagueDate.FromCode(formatType))
                        {
                            case VagueDateTypes.StartDate: // "D"
                                if (startDateOk) return startDate.Subtract(BaseDate).Days;
                                break;
                            case VagueDateTypes.StartMonthAndYear: // "O"
                                if (startDateOk) return startDate.Subtract(BaseDate).Days;
                                break;
                            case VagueDateTypes.StartYear: // "Y"
                                if (startDateOk)
                                    return new DateTime(startDate.Year, 1, 1).Subtract(BaseDate).Days;
                                else if (Int32.TryParse(startDateString, out year))
                                    return new DateTime(year, 1, 1).Subtract(BaseDate).Days;
                                break;
                            case VagueDateTypes.StartSeason: // "P"
                                if (!String.IsNullOrEmpty(startDateString))
                                {
                                    string[] splitArray = Regex.Split(startDateString, @"\s+");
                                    if ((splitArray.Length != 2) || !Int32.TryParse(splitArray[1], out year))
                                        return DateUnknown;
                                    var q = SeasonNames.Where(s => s.ToLower() == splitArray[0].ToLower());
                                    if (q.Count() > 0)
                                    {
                                        string season = q.ElementAt(0);
                                        int seasonIx = System.Array.IndexOf(SeasonNames, season);
                                        if (seasonIx != -1)
                                            return SeasonStart(splitArray[0], year).Subtract(BaseDate).Days;
                                    }
                                }
                                break;
                            case VagueDateTypes.Unknown: // "U"
                            default:
                                return DateUnknown;
                        }
                        break;

                    case DateType.End:
                        if (dateType.Length > 1)
                            formatType = dateType.Substring(dateType.Length - 1, 1);
                        else if (dateType.Length == 1)
                            formatType = dateType.Substring(0, 1);
                        switch (VagueDate.FromCode(formatType))
                        {
                            // Return the exact day
                            case VagueDateTypes.StartDate: // "D"
                                if (endDateOk) return endDate.Subtract(BaseDate).Days;
                                break;

                            // Return the end of the month
                            case VagueDateTypes.StartMonthAndYear: // "O"
                                //if (endDateOk) return endDate.Subtract(BaseDate).Days;
                                if (!String.IsNullOrEmpty(endDateString))
                                {
                                    string[] splitArray = Regex.Split(endDateString, @"\s+");
                                    if ((splitArray.Length != 2) || !Int32.TryParse(splitArray[1], out year))
                                        return DateUnknown;
                                    var q = dtFormatInfo.MonthNames.Where(s => s.ToLower() == splitArray[0].ToLower());
                                    if (q.Count() > 0)
                                    {
                                        string month = q.ElementAt(0);
                                        int monthIx = System.Array.IndexOf(dtFormatInfo.MonthNames, month) + 1;
                                        return new DateTime(year, 1, 1).AddMonths(monthIx).AddDays(-1).Subtract(BaseDate).Days;
                                    }
                                }
                                break;

                            // Return the end of the year
                            case VagueDateTypes.StartYear: // "Y"
                                if (endDateOk)
                                    return new DateTime(endDate.Year, 1, 1).Subtract(BaseDate).Days;
                                else if (Int32.TryParse(endDateString, out year))
                                    return new DateTime(year, 12, 31).Subtract(BaseDate).Days;
                                break;

                            // Return the end of the season
                            case VagueDateTypes.StartSeason: // "P"
                                if (!String.IsNullOrEmpty(endDateString))
                                {
                                    string[] splitArray = Regex.Split(endDateString, @"\s+");
                                    if ((splitArray.Length != 2) || !Int32.TryParse(splitArray[1], out year))
                                        return DateUnknown;
                                    var q = SeasonNames.Where(s => s.ToLower() == splitArray[0].ToLower());
                                    if (q.Count() > 0)
                                    {
                                        string season = q.ElementAt(0);
                                        int seasonIx = System.Array.IndexOf(SeasonNames, season);
                                        return SeasonEnd(season, year).Subtract(BaseDate).Days;
                                    }
                                }
                                break;

                            // Return unknown date value
                            case VagueDateTypes.Unknown: // "U"
                            default:
                                return DateUnknown;
                        }
                        break;
                }
            }
            catch { }

            return DateUnknown; // 9999
        }

        /// <summary>
        /// Splits a vague date string into start and end date strings using
        /// Settings.Default.VagueDateDelimiter as the delimiter (user setting).
        /// </summary>
        /// <param name="dateString">Vague date string to be split.</param>
        /// <param name="startDateString">The start date portion of dateString.</param>
        /// <param name="endDateString">The end date portion of dateString.</param>
        /// <returns>True if split succeeds, otherwise false.</returns>
        private static bool SplitDateString(string dateString,
            out string startDateString, out string endDateString, out bool delimiterFound)
        {
            startDateString = null;
            endDateString = null;
            delimiterFound = false;
            if (String.IsNullOrEmpty(dateString)) return false;
            string[] a = Regex.Split(dateString, @"\s*" + Delimiter + @"\s*");

            startDateString = a[0];
            if (a.Length == 2) endDateString = a[1];
            if (dateString.Contains(Delimiter)) delimiterFound = true;
            return a.Length <= 2;
        }

        /// <summary>
        /// Formats a season string from a vague date portion.
        /// </summary>
        /// <param name="date">Start or end date portion of input vague date.</param>
        /// <returns>Formatted date string. For winter, year will be that of start of season.</returns>
        private static string SeasonString(DateTime date)
        {
            int season = SeasonIx(date);

            return String.Format("{0}{1}", SeasonNames[SeasonIx(date)], date.Year == 9999 ||
                date.Year == -1 ? String.Empty : " " + (season == 3 && date == SeasonEnd(SeasonNames[3], date.Year - 1) ?
                date.Year - 1 : date.Year).ToString());
        }

        private static int SeasonIx(DateTime date)
        {
            int doy = date.DayOfYear -
                Convert.ToInt32((DateTime.IsLeapYear(date.Year)) && date.DayOfYear > 59);

            if (doy < 80 || doy >= 355)
                return 3; // winter
            else if (doy >= 80 && doy < 172)
                return 0; // spring
            else if (doy >= 172 && doy < 266)
                return 1; // summer
            else
                return 2; // autumn
        }

        private static DateTime SeasonStart(string season, int year)
        {
            // If no season return first day of year
            if (String.IsNullOrEmpty(season))
            {
                if (year > 0)
                    return new DateTime(year, 1, 1);
                else
                    return DateTime.MinValue;
            }

            // Find the season index number (and check it is valid)
            int ix = System.Array.IndexOf<string>(SeasonNames, season);

            // Return the first day of the year plus the number of
            // days to the first day the season
            switch (ix)
            {
                case 0:
                    return new DateTime(year, 1, 1).AddDays(80);
                case 1:
                    return new DateTime(year, 1, 1).AddDays(172);
                case 2:
                    return new DateTime(year, 1, 1).AddDays(266);
                case 3:
                    return new DateTime(year, 1, 1).AddDays(355);
                default:
                    return DateTime.MinValue;
            }
        }

        private static DateTime SeasonEnd(string season, int year)
        {
            // If no season return first day of year
            if (String.IsNullOrEmpty(season))
            {
                if (year > 0)
                    return new DateTime(year, 1, 1);
                else
                    return DateTime.MinValue;
            }

            // Find the season index number (and check it is valid)
            int ix = System.Array.IndexOf<string>(SeasonNames, season);

            // Return the first day of the year plus the number of
            // days to the last day the season
            switch (ix)
            {
                case 0:
                    return new DateTime(year, 1, 1).AddDays(170);
                case 1:
                    return new DateTime(year, 1, 1).AddDays(264);
                case 2:
                    return new DateTime(year, 1, 1).AddDays(353);
                case 3:
                    return new DateTime(year + 1, 1, 1).AddDays(78);
                default:
                    return DateTime.MinValue;
            }
        }

        private static DateTime MonthEnd(string season, int month, int year)
        {
            // If not season return first day of year
            if (String.IsNullOrEmpty(season))
            {
                if (year > 0)
                    return new DateTime(year, 1, 1);
                else
                    return DateTime.MinValue;
            }

            // Return the first day of the year plus the number of
            // days to the last day the month
            return new DateTime(year, 1, 1).AddMonths(month).AddDays(-1);
        }
    
}
}
