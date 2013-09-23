using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Media;

namespace HLU.UI.UserControls
{
    static class FindControls
    {
        public static IEnumerable<T> FindLogicalChildren<T>(DependencyObject depObj) where T : DependencyObject
        {
            if (depObj != null)
            {
                foreach (object c in LogicalTreeHelper.GetChildren(depObj))
                {
                    DependencyObject child = c as DependencyObject;
                    if ((child != null) && (child is T))
                    {
                        yield return (T)child;
                    }

                    foreach (T childOfChild in FindLogicalChildren<T>(child))
                    {
                        yield return childOfChild;
                    }
                }
            }
        }

        public static IEnumerable<T> FindVisualChildren<T>(DependencyObject depObj) where T : DependencyObject
        {
            if (depObj != null)
            {
                for (int i = 0; i < VisualTreeHelper.GetChildrenCount(depObj); i++)
                {
                    DependencyObject child = VisualTreeHelper.GetChild(depObj, i);
                    if (child != null && child is T)
                    {
                        yield return (T)child;
                    }

                    foreach (T childOfChild in FindVisualChildren<T>(child))
                    {
                        yield return childOfChild;
                    }
                }
            }
        }

        public static DependencyObject GetChildren(this DependencyObject reference, Type childType,
            ref List<DependencyObject> childrenList)
        {
            DependencyObject foundChild = null;
            if (childrenList == null) childrenList = new List<DependencyObject>();
            if (reference != null)
            {
                int childrenCount = VisualTreeHelper.GetChildrenCount(reference);
                for (int i = 0; i < childrenCount; i++)
                {
                    var child = VisualTreeHelper.GetChild(reference, i);
                    if (child.GetType() != childType)
                    {
                        foundChild = GetChildren(child, childType, ref childrenList);
                    }
                    else
                    {
                        childrenList.Add(child);
                    }
                }
            }
            return foundChild;
        }

        public static DependencyObject FindChild(this DependencyObject reference, string childName, Type childType)
        {
            DependencyObject foundChild = null;
            if (reference != null)
            {
                int childrenCount = VisualTreeHelper.GetChildrenCount(reference);
                for (int i = 0; i < childrenCount; i++)
                {
                    var child = VisualTreeHelper.GetChild(reference, i);
                    if (child.GetType() != childType)
                    {
                        foundChild = FindChild(child, childName, childType);
                    }
                    else if (!string.IsNullOrEmpty(childName))
                    {
                        var frameworkElement = child as FrameworkElement;
                        if (frameworkElement != null && frameworkElement.Name == childName)
                        {
                            foundChild = child;
                            break;
                        }
                    }
                    else
                    {
                        foundChild = child;
                        break;
                    }
                }
            }
            return foundChild;
        }
    }
}
