// HLUTool is used to view and maintain habitat and land use GIS data.
// Copyright © 2011 Hampshire Biodiversity Information Centre
// Copyright © 2014 Sussex Biodiversity Record Centre
// Copyright © 2016 Thames Valley Environmental Records Centre
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
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Windows;
using HLU.Data;

namespace HLU.UI.ViewModel
{
    public class WindowMainCopySwitches : INotifyPropertyChanged
    {
        #region Fields

        #region Switches

        private bool _copyIncidPrimary;
        private bool _copyIncidSecondaryHabitats;
        private bool _copyIncidLegacyHabitat;
        private bool _copyIncidBapHabitatsUser;
        private bool _copyIncidGeneralComments;
        private bool _copyIncidBoundaryBaseMap;
        private bool _copyIncidDigitisationBaseMap;
        private bool _copyIncidSiteRef;
        private bool _copyIncidSiteName;
        private bool _copyIncidCondition;
        private bool _copyIncidConditionQualifier;
        private bool _copyIncidConditionDate;
        private bool _copyIncidQualityDetermination;
        private bool _copyIncidQualityInterpretation;
        private bool _copyIncidQualityComments;
        private bool _copyIncidSource1Id;
        private bool _copyIncidSource1Date;
        private bool _copyIncidSource1HabitatClass;
        private bool _copyIncidSource1HabitatType;
        private bool _copyIncidSource1BoundaryImportance;
        private bool _copyIncidSource1HabitatImportance;
        private bool _copyIncidSource2Id;
        private bool _copyIncidSource2Date;
        private bool _copyIncidSource2HabitatClass;
        private bool _copyIncidSource2HabitatType;
        private bool _copyIncidSource2BoundaryImportance;
        private bool _copyIncidSource2HabitatImportance;
        private bool _copyIncidSource3Id;
        private bool _copyIncidSource3Date;
        private bool _copyIncidSource3HabitatClass;
        private bool _copyIncidSource3HabitatType;
        private bool _copyIncidSource3BoundaryImportance;
        private bool _copyIncidSource3HabitatImportance;

        #endregion

        #region Values

        private string _incidPrimary;
        private ObservableCollection<SecondaryHabitat> _incidSecondaryHabitats;
        private string _incidLegacyHabitat;
        private ObservableCollection<BapEnvironment> _incidBapHabitatsUser;
        private string _incidGeneralComments;
        private string _incidBoundaryBaseMap;
        private string _incidDigitisationBaseMap;
        private string _incidSiteRef;
        private string _incidSiteName;
        private string _incidCondition;
        private string _incidConditionQualifier;
        private Date.VagueDateInstance _incidConditionDate;
        private string _incidQualityDetermination;
        private string _incidQualityInterpretation;
        private string _incidQualityComments;
        private Nullable<int> _incidSource1Id;
        private Date.VagueDateInstance _incidSource1Date;
        private string _incidSource1HabitatClass;
        private string _incidSource1HabitatType;
        private string _incidSource1BoundaryImportance;
        private string _incidSource1HabitatImportance;
        private Nullable<int> _incidSource2Id;
        private Date.VagueDateInstance _incidSource2Date;
        private string _incidSource2HabitatClass;
        private string _incidSource2HabitatType;
        private string _incidSource2BoundaryImportance;
        private string _incidSource2HabitatImportance;
        private Nullable<int> _incidSource3Id;
        private Date.VagueDateInstance _incidSource3Date;
        private string _incidSource3HabitatClass;
        private string _incidSource3HabitatType;
        private string _incidSource3BoundaryImportance;
        private string _incidSource3HabitatImportance;

        #endregion

        #endregion

        #region Properties

        #region Switches

        public bool CopyIncidPrimary
        {
            get { return _copyIncidPrimary; }
            set
            {
                _copyIncidPrimary = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("CopyIncidPrimary"));
            }
        }

        public bool CopyIncidSecondaryHabitats
        {
            get { return _copyIncidSecondaryHabitats; }
            set
            {
                _copyIncidSecondaryHabitats = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("CopyIncidSecondaryHabitats"));
            }
        }

        //---------------------------------------------------------------------
        // CHANGED: CR44 (Editable Legacy Habitat field)
        // Make the legacy habitat field editable in the interface.
        public bool CopyIncidLegacyHabitat
        {
            get { return _copyIncidLegacyHabitat; }
            set
            {
                _copyIncidLegacyHabitat = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("CopyIncidLegacyHabitat"));
            }
        }
        //---------------------------------------------------------------------

        public bool CopyIncidBapHabitatsUser
        {
            get { return _copyIncidBapHabitatsUser; }
            set
            {
                _copyIncidBapHabitatsUser = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("CopyIncidBapHabitatsUser"));
            }
        }

        public bool CopyIncidGeneralComments
        {
            get { return _copyIncidGeneralComments; }
            set
            {
                _copyIncidGeneralComments = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("CopyIncidGeneralComments"));
            }
        }

        public bool CopyIncidBoundaryBaseMap
        {
            get { return _copyIncidBoundaryBaseMap; }
            set
            {
                _copyIncidBoundaryBaseMap = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("CopyIncidBoundaryBaseMap"));
            }
        }

        public bool CopyIncidDigitisationBaseMap
        {
            get { return _copyIncidDigitisationBaseMap; }
            set
            {
                _copyIncidDigitisationBaseMap = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("CopyIncidDigitisationBaseMaps"));
            }
        }

        //---------------------------------------------------------------------
        // CHANGED: CR37 (Site reference and site name)
        // Display the site reference with the site name in the interface.
        public bool CopyIncidSiteRef
        {
            get { return _copyIncidSiteRef; }
            set
            {
                _copyIncidSiteRef = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("CopyIncidSiteRef"));
            }
        }
        //---------------------------------------------------------------------

        public bool CopyIncidSiteName
        {
            get { return _copyIncidSiteName; }
            set
            {
                _copyIncidSiteName = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("CopyIncidSiteName"));
            }
        }

        public bool CopyIncidCondition
        {
            get { return _copyIncidCondition; }
            set
            {
                _copyIncidCondition = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("CopyIncidCondition"));
            }
        }

        public bool CopyIncidConditionQualifier
        {
            get { return _copyIncidConditionQualifier; }
            set
            {
                _copyIncidConditionQualifier = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("CopyIncidConditionQualifier"));
            }
        }

        public bool CopyIncidConditionDate
        {
            get { return _copyIncidConditionDate; }
            set
            {
                _copyIncidConditionDate = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("CopyIncidConditionDate"));
            }
        }

        public bool CopyIncidQualityDetermination
        {
            get { return _copyIncidQualityDetermination; }
            set
            {
                _copyIncidQualityDetermination = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("CopyIncidQualityDetermination"));
            }
        }

        public bool CopyIncidQualityInterpretation
        {
            get { return _copyIncidQualityInterpretation; }
            set
            {
                _copyIncidQualityInterpretation = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("CopyIncidQualityInterpretation"));
            }
        }

        public bool CopyIncidQualityComments
        {
            get { return _copyIncidQualityComments; }
            set
            {
                _copyIncidQualityComments = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("CopyIncidQualityComments"));
            }
        }

        public bool CopyIncidSource1Id
        {
            get { return _copyIncidSource1Id; }
            set
            {
                _copyIncidSource1Id = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("CopyIncidSource1Id"));
            }
        }

        public bool CopyIncidSource1Date
        {
            get { return _copyIncidSource1Date; }
            set
            {
                _copyIncidSource1Date = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("CopyIncidSource1Date"));
            }
        }

        public bool CopyIncidSource1HabitatClass
        {
            get { return _copyIncidSource1HabitatClass; }
            set
            {
                _copyIncidSource1HabitatClass = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("CopyIncidSource1HabitatClass"));
            }
        }

        public bool CopyIncidSource1HabitatType
        {
            get { return _copyIncidSource1HabitatType; }
            set
            {
                _copyIncidSource1HabitatType = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("CopyIncidSource1HabitatType"));
            }
        }

        public bool CopyIncidSource1BoundaryImportance
        {
            get { return _copyIncidSource1BoundaryImportance; }
            set
            {
                _copyIncidSource1BoundaryImportance = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("CopyIncidSource1BoundaryImportance"));
            }
        }

        public bool CopyIncidSource1HabitatImportance
        {
            get { return _copyIncidSource1HabitatImportance; }
            set
            {
                _copyIncidSource1HabitatImportance = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("CopyIncidSource1HabitatImportance"));
            }
        }

        public bool CopyIncidSource2Id
        {
            get { return _copyIncidSource2Id; }
            set
            {
                _copyIncidSource2Id = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("CopyIncidSource2Id"));
            }
        }

        public bool CopyIncidSource2Date
        {
            get { return _copyIncidSource2Date; }
            set
            {
                _copyIncidSource2Date = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("CopyIncidSource2Date"));
            }
        }

        public bool CopyIncidSource2HabitatClass
        {
            get { return _copyIncidSource2HabitatClass; }
            set
            {
                _copyIncidSource2HabitatClass = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("CopyIncidSource2HabitatClass"));
            }
        }

        public bool CopyIncidSource2HabitatType
        {
            get { return _copyIncidSource2HabitatType; }
            set
            {
                _copyIncidSource2HabitatType = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("CopyIncidSource2HabitatType"));
            }
        }

        public bool CopyIncidSource2BoundaryImportance
        {
            get { return _copyIncidSource2BoundaryImportance; }
            set
            {
                _copyIncidSource2BoundaryImportance = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("CopyIncidSource2BoundaryImportance"));
            }
        }

        public bool CopyIncidSource2HabitatImportance
        {
            get { return _copyIncidSource2HabitatImportance; }
            set
            {
                _copyIncidSource2HabitatImportance = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("CopyIncidSource2HabitatImportance"));
            }
        }

        public bool CopyIncidSource3Id
        {
            get { return _copyIncidSource3Id; }
            set
            {
                _copyIncidSource3Id = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("CopyIncidSource3Id"));
            }
        }

        public bool CopyIncidSource3Date
        {
            get { return _copyIncidSource3Date; }
            set
            {
                _copyIncidSource3Date = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("CopyIncidSource3Date"));
            }
        }

        public bool CopyIncidSource3HabitatClass
        {
            get { return _copyIncidSource3HabitatClass; }
            set
            {
                _copyIncidSource3HabitatClass = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("CopyIncidSource3HabitatClass"));
            }
        }

        public bool CopyIncidSource3HabitatType
        {
            get { return _copyIncidSource3HabitatType; }
            set
            {
                _copyIncidSource3HabitatType = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("CopyIncidSource3HabitatType"));
            }
        }

        public bool CopyIncidSource3BoundaryImportance
        {
            get { return _copyIncidSource3BoundaryImportance; }
            set
            {
                _copyIncidSource3BoundaryImportance = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("CopyIncidSource3BoundaryImportance"));
            }
        }

        public bool CopyIncidSource3HabitatImportance
        {
            get { return _copyIncidSource3HabitatImportance; }
            set
            {
                _copyIncidSource3HabitatImportance = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("CopyIncidSource3HabitatImportance"));
            }
        }

        #endregion

        #region Values

        public string IncidPrimary
        {
            get { return _copyIncidPrimary ? _incidPrimary : null; }
            set
            {
                _incidPrimary = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("IncidPrimary"));
            }
        }

        public ObservableCollection<SecondaryHabitat> IncidSecondaryHabitats
        {
            get { return _copyIncidSecondaryHabitats ? _incidSecondaryHabitats : null; }
            set
            {
                ObservableCollection<SecondaryHabitat> newSH = new ObservableCollection<SecondaryHabitat>();
                foreach (var sh in value)
                {
                    SecondaryHabitat shClone = (SecondaryHabitat)sh.Clone();
                    newSH.Add(shClone);
                }
                _incidSecondaryHabitats = newSH;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("IncidSecondaryHabitats"));
            }
        }

        //---------------------------------------------------------------------
        // CHANGED: CR44 (Editable Legacy Habitat field)
        // Make the legacy habitat field editable in the interface.
        public string IncidLegacyHabitat
        {
            get { return _copyIncidLegacyHabitat ? _incidLegacyHabitat : null; }
            set
            {
                _incidLegacyHabitat = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("IncidLegacyHabitat"));
            }
        }
        //---------------------------------------------------------------------

        public ObservableCollection<BapEnvironment> IncidBapHabitatsUser
        {
            get { return _copyIncidBapHabitatsUser ? _incidBapHabitatsUser : null; }
            set
            {
                ObservableCollection<BapEnvironment> newBH = new ObservableCollection<BapEnvironment>();
                foreach (var bh in value)
                {
                    BapEnvironment bhClone = (BapEnvironment)bh.Clone();
                    newBH.Add(bhClone);
                }
                _incidBapHabitatsUser = newBH;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("IncidBapHabitatsUser"));
 }
        }

        public string IncidGeneralComments
        {
            get { return _incidGeneralComments; }
            set
            {
                _incidGeneralComments = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("IncidGeneralComments"));
            }
        }

        public string IncidBoundaryBaseMap
        {
            get { return _incidBoundaryBaseMap; }
            set
            {
                _incidBoundaryBaseMap = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("IncidBoundaryBaseMap"));
            }
        }

        public string IncidDigitisationBaseMap
        {
            get { return _incidDigitisationBaseMap; }
            set
            {
                _incidDigitisationBaseMap = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("IncidDigitisationBaseMap"));
            }
        }

        //---------------------------------------------------------------------
        // CHANGED: CR37 (Site reference and site name)
        // Display the site reference with the site name in the interface.
        public string IncidSiteRef
        {
            get { return _incidSiteRef; }
            set
            {
                _incidSiteRef = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("IncidSiteRef"));
            }
        }
        //---------------------------------------------------------------------

        public string IncidSiteName
        {
            get { return _incidSiteName; }
            set
            {
                _incidSiteName = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("IncidSiteName"));
            }
        }

        public string IncidCondition
        {
            get { return _copyIncidCondition ? _incidCondition : null; }
            set
            {
                _incidCondition = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("IncidCondition"));
            }
        }

        public string IncidConditionQualifier
        {
            get { return _copyIncidConditionQualifier ? _incidConditionQualifier : null; }
            set
            {
                _incidConditionQualifier = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("IncidConditionQualifier"));
            }
        }

        public Date.VagueDateInstance IncidConditionDate
        {
            get { return _copyIncidConditionDate ? _incidConditionDate : null; }
            set
            {
                _incidConditionDate = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("IncidConditionDate"));
            }
        }

        public string IncidQualityDetermination
        {
            get { return _copyIncidQualityDetermination ? _incidQualityDetermination : null; }
            set
            {
                _incidQualityDetermination = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("IncidQualityDetermination"));
            }
        }

        public string IncidQualityInterpretation
        {
            get { return _copyIncidQualityInterpretation ? _incidQualityInterpretation : null; }
            set
            {
                _incidQualityInterpretation = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("IncidQualityInterpretation"));
            }
        }

        public string IncidQualityComments
        {
            get { return _copyIncidQualityComments ? _incidQualityComments : null; }
            set
            {
                _incidQualityComments = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("IncidQualityComments"));
            }
        }

        public Nullable<int> IncidSource1Id
        {
            get { return _copyIncidSource1Id ? _incidSource1Id : null; }
            set
            {
                _incidSource1Id = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("IncidSource1Id"));
            }
        }

        public Date.VagueDateInstance IncidSource1Date
        {
            get { return _copyIncidSource1Date ? _incidSource1Date : null; }
            set
            {
                _incidSource1Date = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("IncidSource1Date"));
            }
        }

        public string IncidSource1HabitatClass
        {
            get { return _copyIncidSource1HabitatClass ? _incidSource1HabitatClass : null; }
            set
            {
                _incidSource1HabitatClass = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("IncidSource1HabitatClass"));
            }
        }

        public string IncidSource1HabitatType
        {
            get { return _copyIncidSource1HabitatType ? _incidSource1HabitatType : null; }
            set
            {
                _incidSource1HabitatType = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("IncidSource1HabitatType"));
            }
        }

        public string IncidSource1BoundaryImportance
        {
            get { return _copyIncidSource1BoundaryImportance ? _incidSource1BoundaryImportance : null; }
            set
            {
                _incidSource1BoundaryImportance = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("IncidSource1BoundaryImportance"));
            }
        }

        public string IncidSource1HabitatImportance
        {
            get { return _copyIncidSource1HabitatImportance ? _incidSource1HabitatImportance : null; }
            set
            {
                _incidSource1HabitatImportance = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("IncidSource1HabitatImportance"));
            }
        }

        public Nullable<int> IncidSource2Id
        {
            get { return _copyIncidSource2Id ? _incidSource2Id : null; }
            set
            {
                _incidSource2Id = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("IncidSource2Id"));
            }
        }

        public Date.VagueDateInstance IncidSource2Date
        {
            get { return _copyIncidSource2Date ? _incidSource2Date : null; }
            set
            {
                _incidSource2Date = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("IncidSource2Date"));
            }
        }

        public string IncidSource2HabitatClass
        {
            get { return _copyIncidSource2HabitatClass ? _incidSource2HabitatClass : null; }
            set
            {
                _incidSource2HabitatClass = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("IncidSource2HabitatClass"));
            }
        }

        public string IncidSource2HabitatType
        {
            get { return _copyIncidSource2HabitatType ? _incidSource2HabitatType : null; }
            set
            {
                _incidSource2HabitatType = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("IncidSource2HabitatType"));
            }
        }

        public string IncidSource2BoundaryImportance
        {
            get { return _copyIncidSource2BoundaryImportance ? _incidSource2BoundaryImportance : null; }
            set
            {
                _incidSource2BoundaryImportance = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("IncidSource2BoundaryImportance"));
            }
        }

        public string IncidSource2HabitatImportance
        {
            get { return _copyIncidSource2HabitatImportance ? _incidSource2HabitatImportance : null; }
            set
            {
                _incidSource2HabitatImportance = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("IncidSource2HabitatImportance"));
            }
        }

        public Nullable<int> IncidSource3Id
        {
            get { return _copyIncidSource3Id ? _incidSource3Id : null; }
            set
            {
                _incidSource3Id = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("IncidSource3Id"));
            }
        }

        public Date.VagueDateInstance IncidSource3Date
        {
            get { return _copyIncidSource3Date ? _incidSource3Date : null; }
            set
            {
                _incidSource3Date = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("IncidSource3Date"));
            }
        }

        public string IncidSource3HabitatClass
        {
            get { return _copyIncidSource3HabitatClass ? _incidSource3HabitatClass : null; }
            set
            {
                _incidSource3HabitatClass = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("IncidSource3HabitatClass"));
            }
        }

        public string IncidSource3HabitatType
        {
            get { return _copyIncidSource3HabitatType ? _incidSource3HabitatType : null; }
            set
            {
                _incidSource3HabitatType = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("IncidSource3HabitatType"));
            }
        }

        public string IncidSource3BoundaryImportance
        {
            get { return _copyIncidSource3BoundaryImportance ? _incidSource3BoundaryImportance : null; }
            set
            {
                _incidSource3BoundaryImportance = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("IncidSource3BoundaryImportance"));
            }
        }

        public string IncidSource3HabitatImportance
        {
            get { return _copyIncidSource3HabitatImportance ? _incidSource3HabitatImportance : null; }
            set
            {
                _incidSource3HabitatImportance = value;
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("IncidSource3HabitatImportance"));
            }
        }

        #endregion

        #endregion

        #region Copy/Paste

        internal void CopyValues(ViewModelWindowMain vmMain)
        {
            List<string> errorProps = new List<string>();
            string propertyName = null;
            object copyValue = null;

            foreach (PropertyInfo pi in this.GetType().GetProperties().Where(p => p.Name.StartsWith("Copy")))
            {
                try
                {
                    propertyName = pi.Name.Remove(0, 4);
                    PropertyInfo copyPI = typeof(WindowMainCopySwitches).GetProperty(propertyName);
                    PropertyInfo valuePI = vmMain.GetType().GetProperty(propertyName);

                    if ((copyPI != null) && (valuePI != null) && (copyPI.PropertyType == valuePI.PropertyType))
                    {
                        if ((bool)pi.GetValue(this, null))
                            copyValue = valuePI.GetValue(vmMain, null);
                        else
                            copyValue = GetDefault(copyPI.PropertyType);

                        // Ensure observable collection properties are set to null correctly
                        if ((propertyName == "IncidSecondaryHabitats") && (copyValue == null))
                            copyPI.SetValue(this, new ObservableCollection<SecondaryHabitat>(), null);
                        else if ((propertyName == "IncidBapHabitatsUser") && (copyValue == null))
                            copyPI.SetValue(this, new ObservableCollection<BapEnvironment>(), null);
                        else
                            copyPI.SetValue(this, copyValue, null);
                    }
                    else
                    {
                        errorProps.Add(propertyName);
                    }
                }
                catch { errorProps.Add(propertyName); }
            }

            if (errorProps.Count > 0)
                MessageBox.Show(String.Format("Unable to copy {0}{1}.", errorProps.Count > 1 ?
                    String.Format("the following {0} fields: \n\n", errorProps.Count) : "field ",
                    String.Join(", ", errorProps.ToArray())), "HLU: Copy Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);

            vmMain.OnPropertyChanged("CanPaste");
        }

        internal void PasteValues(ViewModelWindowMain vmMain)
        {
            List<string> errorProps = new List<string>();
            string propertyName = null;

            foreach (PropertyInfo pi in this.GetType().GetProperties().Where(p => p.Name.StartsWith("Copy")))
            {
                try
                {
                    propertyName = pi.Name.Remove(0, 4);

                    PropertyInfo copyPI = this.GetType().GetProperty(propertyName);
                    PropertyInfo valuePI = vmMain.GetType().GetProperty(propertyName);

                    if ((copyPI != null) && (valuePI != null) && (copyPI.PropertyType == valuePI.PropertyType))
                    {
                        if ((bool)pi.GetValue(this, null))
                        {
                            vmMain.Pasting = true;
                            
                            valuePI.SetValue(vmMain, copyPI.GetValue(this, null), null);

                            // Refresh the secondary habitat table if they have been pasted
                            if (propertyName == "IncidSecondaryHabitats")
                                vmMain.RefreshSecondaryHabitats();

                            vmMain.Pasting = false;
                        }
                    }
                    else
                    {
                        errorProps.Add(propertyName);
                    }
                }
                catch { errorProps.Add(propertyName); }
                finally { vmMain.Pasting = false; }
            }

            if (errorProps.Count > 0)
                MessageBox.Show(String.Format("Unable to paste {0}{1}", errorProps.Count > 1 ?
                    String.Format("the following {0} fields: \n\n", errorProps.Count) : "field ",
                    String.Join(", ", errorProps.ToArray())), "HLU: Paste Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);

            // Refresh all the controls
            vmMain.RefreshAll();

        }

        private object GetDefault(Type targetType)
        {
            return targetType.IsValueType ? Activator.CreateInstance(targetType) : null;
        }

        #endregion

        #region INotifyPropertyChanged Members

        public event PropertyChangedEventHandler PropertyChanged;

        #endregion
    }
}
