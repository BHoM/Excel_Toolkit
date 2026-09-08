/*
 * This file is part of the Buildings and Habitats object Model (BHoM)
 * Copyright (c) 2015 - 2026, the respective contributors. All rights reserved.
 *
 * Each contributor holds copyright over their respective contributions.
 * The project versioning (Git) records all such contribution source information.
 *
 *
 * The BHoM is free software: you can redistribute it and/or modify
 * it under the terms of the GNU Lesser General Public License as published by
 * the Free Software Foundation, either version 3.0 of the License, or
 * (at your option) any later version.
 *
 * The BHoM is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 * GNU Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public License
 * along with this code. If not, see <https://www.gnu.org/licenses/lgpl-3.0.html>.
 */

using BH.Engine.Reflection;
using BH.oM.Adapters.Excel;
using BH.oM.Base.Attributes;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;

namespace BH.Engine.Excel
{
    public static partial class Convert
    {
        /***************************************************/
        /**** Public Methods                            ****/
        /***************************************************/

        public static List<TableRow> ToTableRows(this IEnumerable<object> objects, ExcelPushConfig config = null)
        {
            if (config == null)
            {
                BH.Engine.Base.Compute.RecordNote("Default push config has been used to convert objects to table rows.");
                config = new ExcelPushConfig();
            }

            return objects.ToTableRows(config.ObjectProperties, config.PropertiesToIgnore, config.GoDeepInProperties, config.TransposeObjectTable, config.IncludePropertyNames);
        }

        /***************************************************/

        [Description("Converts a collection of objects to a list of table rows.")]
        [Input("objects", "Collection of objects to convert.")]
        [Input("properties", "List of properties to include in the table.")]
        [Input("propertiesToIgnore", "List of properties to ignore in the table.")]
        [Input("goDeep", "If true, will go deep into the object properties.")]
        [Input("transposeTable", "If true, will transpose the table.")]
        [Input("showPropertyNames", "If true, will show the property names in the first row.")]
        [Output("tableRows", "List of table rows created based on the input objects.")]
        public static List<TableRow> ToTableRows(this IEnumerable<object> objects, List<string> properties, List<string> propertiesToIgnore, bool goDeep = false, bool transposeTable = false, bool showPropertyNames = true)
        {
            // Get the property dictionary for the object
            List<Dictionary<string, object>> props = GetPropertyDictionaries(objects, goDeep);
            if (props.Count < 1)
                return new List<TableRow>();

            // Create the list of keys
            List<string> keys = new List<string>();
            if (properties?.Count == 0)
                keys = props.SelectMany(x => x.Keys).Distinct().ToList();
            else
                keys = properties.ToList();

            if (propertiesToIgnore?.Count > 0)
            {
                if (properties?.Count == 0)
                    propertiesToIgnore = propertiesToIgnore.Except(properties).ToList();
                keys = keys.Except(propertiesToIgnore).ToList();
            }

            // Get the exploded table
            List<List<object>> result = new List<List<object>>();
            if (showPropertyNames)
                result.Add(keys.ToList<object>());

            for (int i = 0; i < props.Count; i++)
                result.Add(keys.Select(k => props[i].ContainsKey(k) ? props[i][k] : null).ToList());

            if (transposeTable)
            {
                result = result.SelectMany(row => row.Select((value, index) => new { value, index }))
                    .GroupBy(cell => cell.index, cell => cell.value)
                    .Select(g => g.ToList()).ToList();
            }

            return result.Select(x => new TableRow { Content = x }).ToList();
        }


        /***************************************************/
        /**** Private Methods                           ****/
        /***************************************************/

        private static List<Dictionary<string, object>> GetPropertyDictionaries(IEnumerable<object> objs, bool goDeep = false)
        {
            //Get the property dictionary for the object
            List<Dictionary<string, object>> props = new List<Dictionary<string, object>>();
            foreach (object obj in objs)
            {
                if (obj is IEnumerable && !(obj is string))
                {
                    props.AddRange(GetPropertyDictionaries((obj as IEnumerable).Cast<object>().ToList(), goDeep));
                }
                else
                {
                    Dictionary<string, object> dict = new Dictionary<string, object>();
                    GetPropertyDictionary(ref dict, obj, goDeep);
                    props.Add(dict);
                }
            }

            return props;
        }

        /*******************************************/

        private static void GetPropertyDictionary(ref Dictionary<string, object> dict, object obj, bool goDeep = false, string parentType = "")
        {
            if (obj == null)
            {
                return;
            }
            else if (obj.GetType().IsPrimitive || obj is string || obj is Guid || obj is Enum)
            {
                if (obj is Guid)
                    obj = obj.ToString(); // TODO: Temporary fix to cover for a bug in IToText() from the base engine when used on Guids. Remove when fix in the base engine.

                string key = parentType.Length > 0 ? parentType : "Value";
                dict[key] = obj;
                return;
            }
            else
            {
                foreach (KeyValuePair<string, object> kvp in obj.PropertyDictionary())
                {
                    string key = (parentType.Length > 0) ? parentType + "." + kvp.Key : kvp.Key;
                    if (goDeep)
                        GetPropertyDictionary(ref dict, kvp.Value, true, key);
                    else
                        dict[key] = kvp.Value;
                }
            }
        }

        /*******************************************/
    }
}
