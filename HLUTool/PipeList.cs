// HLUTool is used to view and maintain habitat and land use GIS data.
// Copyright © 2011 Hampshire Biodiversity Information Centre
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
using HLU.Properties;

namespace HLU
{
    public class PipeList
    {
        static string _transmissionEnd = Settings.Default.PipeTransmissionEnd;
        static string _stringContinue = Settings.Default.PipeStringContinue.ToString();
        static int _maxReadBytes = Settings.Default.PipeMaxReadBytes;
        List<string> _pipeList;

        public PipeList(string list, char[] delimiters)
        {
            string[] inList = list.Split(delimiters);
            _pipeList = new List<string>();
            _pipeList.AddRange(inList);
        }

        public PipeList(string[] inList)
        {
            _pipeList = new List<string>();
            _pipeList.AddRange(inList);
        }

        public PipeList(List<string> inList)
        {
            _pipeList = inList;
        }

        public List<string> List
        {
            get
            {
                _pipeList = SplitLongPipeStrings(_pipeList);
                if (_pipeList[_pipeList.Count-1] != _transmissionEnd) _pipeList.Add(_transmissionEnd);
                return _pipeList;
            }
        }

        private List<string> SplitLongPipeStrings(List<string> inList)
        {
            List<String> outList = new List<String>();
            for (int i = 0; i < inList.Count; i++)
            {
                string s = inList[i];
                if (s.Length < _maxReadBytes)
                {
                    outList.Add(s);
                }
                else
                {
                    int limit = s.Length / _maxReadBytes;
                    int remainder = s.Length % _maxReadBytes;
                    for (int j = 0; j < limit; j++)
                    {
                        outList.Add(s.Substring(j * _maxReadBytes, _maxReadBytes));
                        outList.Add(_stringContinue);
                    }
                    if (remainder != 0) outList.Add(s.Substring(s.Length - remainder, remainder));
                }
            }
            return outList;
        }
    }
}
