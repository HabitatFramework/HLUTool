﻿// HLUTool is used to view and maintain habitat and land use GIS data.
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
using System.Runtime.InteropServices;
using MapinfoWrapper;

namespace HLU.GISApplication.MapInfo
{
    [ComVisible(true)]
    public class MapInfoCustomCallback : MapinfoCallback
    {
        public event Action<string> OnMenuItemClick;

        public void MenuItemHandler(string command)
        {
            // Store the event locally to save against a race condition.
            Action<string> menuEvent = OnMenuItemClick;
            if (menuEvent != null)
            {
                // Raise the event.
                menuEvent(command);
            }
        }
    }
}
