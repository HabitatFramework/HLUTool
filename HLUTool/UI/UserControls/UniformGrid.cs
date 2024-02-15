// HLUTool is used to view and maintain habitat and land use GIS data.
// Copyright © 2019 London & South East Record Centres (LaSER)
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
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;

namespace HLU.UI.UserControls
{
    /// <summary>
    /// Enables a custom uniform grid to be used in a user interface
    /// (not currently used).
    /// </summary>
    /// <seealso cref="System.Windows.Controls.Primitives.UniformGrid" />
    public class UniformGrid2 : UniformGrid
    {
        private int _columns;
        private int _rows;

        #region Orientation

        /// <summary>
        /// Orientation Dependency Property
        /// </summary>
        public static readonly DependencyProperty OrientationProperty =
            StackPanel.OrientationProperty.AddOwner(typeof(UniformGrid2),
                new FrameworkPropertyMetadata((Orientation)Orientation.Horizontal,
                    FrameworkPropertyMetadataOptions.AffectsMeasure));

        /// <summary>
        /// Gets or sets the Orientation property.
        /// </summary>
        public Orientation Orientation
        {
            get { return (Orientation)GetValue(OrientationProperty); }
            set { SetValue(OrientationProperty, value); }
        }

        #endregion

        protected override Size ArrangeOverride(Size arrangeSize)
        {
            if (Orientation == Orientation.Horizontal)
                return base.ArrangeOverride(arrangeSize);
            else
                return ArrangeOverrideVertical(arrangeSize);
        }

        protected override Size MeasureOverride(Size constraint)
        {
            if (Orientation == Orientation.Horizontal)
                return base.MeasureOverride(constraint);
            else
                return MeasureOverrideVertical(constraint);
        }

        private Size ArrangeOverrideVertical(Size arrangeSize)
        {
            Rect finalRect = new Rect(0.0, 0.0, arrangeSize.Width / ((double)_columns), arrangeSize.Height / ((double)_rows));
            double height = finalRect.Height;
            double totalHeight = arrangeSize.Height - 1.0;
            foreach (UIElement element in base.InternalChildren)
            {
                element.Arrange(finalRect);
                if (element.Visibility != Visibility.Collapsed)
                {
                    finalRect.Y += height;
                    if (finalRect.Y >= totalHeight)
                    {
                        finalRect.X += finalRect.Width;
                        finalRect.Y = 0.0;
                    }
                }
            }
            return arrangeSize;
        }

        private Size MeasureOverrideVertical(Size constraint)
        {
            UpdateComputedValuesVertical();
            Size availableSize = new Size(constraint.Width / ((double)_columns), constraint.Height / ((double)_rows));
            double width = 0.0;
            double height = 0.0;
            int i = 0;
            int count = base.InternalChildren.Count;
            while (i < count)
            {
                UIElement element = base.InternalChildren[i];
                element.Measure(availableSize);
                Size desiredSize = element.DesiredSize;
                if (width < desiredSize.Width)
                {
                    width = desiredSize.Width;
                }
                if (height < desiredSize.Height)
                {
                    height = desiredSize.Height;
                }
                i++;
            }
            return new Size(width * _columns, height * _rows);
        }

        private void UpdateComputedValuesVertical()
        {
            _columns = Columns;
            _rows = Rows;
            // Ignore FirstColumn property
            FirstColumn = 0;

            if (_rows == 0 || _columns == 0)
            {
                int visibleChildren = 0;
                int i = 0;
                int count = base.InternalChildren.Count;
                while (i < count)
                {
                    UIElement element = base.InternalChildren[i];
                    if (element.Visibility != Visibility.Collapsed)
                    {
                        visibleChildren++;
                    }
                    i++;
                }
                if (visibleChildren == 0)
                {
                    visibleChildren = 1;
                }
                if (_columns == 0)
                {
                    if (_rows > 0)
                    {
                        _columns = (visibleChildren + (_rows - 1)) / _rows;
                    }
                    else
                    {
                        _columns = (int)Math.Sqrt((double)visibleChildren);
                        if ((_columns * _columns) < visibleChildren)
                        {
                            _columns++;
                        }
                        _rows = _columns;
                    }
                }
                else if (_rows == 0)
                {
                    _rows = (visibleChildren + (_columns - 1)) / _columns;
                }
            }
        }
    }
}
