/*
 * This file is part of the Buildings and Habitats object Model (BHoM)
 * Copyright (c) 2015 - 2023, the respective contributors. All rights reserved.
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

using BH.Engine.Adapter;
using BH.Engine.Reflection;
using BH.oM.Adapter;
using BH.oM.Adapters.Excel;
using BH.oM.Base;
using BH.oM.Data.Collections;
using ClosedXML.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace BH.Adapter.Excel
{
    public partial class ExcelAdapter : BHoMAdapter
    {
        /***************************************************/
        /**** Public Overrides                          ****/
        /***************************************************/

        public override List<object> Push(IEnumerable<object> objects, string tag = "", PushType pushType = PushType.AdapterDefault, ActionConfig actionConfig = null)
        {
            if (objects == null || !objects.Any())
            {
                BH.Engine.Base.Compute.RecordError("No objects were provided for Push action.");
                return new List<object>();
            }
            objects = objects.Where(x => x != null).ToList();

            // If unset, set the pushType to AdapterSettings' value (base AdapterSettings default is FullCRUD).
            if (pushType == PushType.AdapterDefault)
                pushType = PushType.DeleteThenCreate;

            // Cast action config to ExcelPushConfig, create new if null.
            ExcelPushConfig config = actionConfig as ExcelPushConfig;
            if (config == null)
            {
                BH.Engine.Base.Compute.RecordNote($"{nameof(ExcelPushConfig)} has not been provided, default config is used.");
                config = new ExcelPushConfig();
            }

            // Make sure that a single type of objects are pushed
            List<Type> objectTypes = objects.Select(x => x.GetType()).Distinct().ToList();
            if (objectTypes.Count != 1)
            {
                string message = "The Excel adapter only allows to push objects of a single type to a table."
                    + "\nRight now you are providing objects of the following types: "
                    + objectTypes.Select(x => x.ToString()).Aggregate((a, b) => a + ", " + b);
                Engine.Base.Compute.RecordError(message);
                return new List<object>();
            }

            // Check if the workbook exists and create it if not.
            string fileName = m_FileSettings.GetFullFileName();
            XLWorkbook workbook;
            if (!File.Exists(fileName))
            {
                if (pushType == PushType.UpdateOnly)
                {
                    BH.Engine.Base.Compute.RecordError($"There is no workbook to update under {fileName}");
                    return new List<object>();
                }

                workbook = new XLWorkbook();
            }
            else
            {
                try
                {
                    workbook = new XLWorkbook(fileName);
                }
                catch (Exception e)
                {
                    BH.Engine.Base.Compute.RecordError($"The existing workbook could not be accessed due to the following error: {e.Message}");
                    return new List<object>();
                }
            }

            // Split the tables into collections to delete, create and update.
            bool success = true;
            string sheetName = config.Worksheet;

            List<TableRow> data = new List<TableRow>();
            if (objects.All(x => x is TableRow))
                data = objects.OfType<TableRow>().ToList();
            else
                data = ToTableRows(objects.ToList(), config.ObjectProperties, config.PropertiesToIgnore, config.GoDeepInProperties, config.TransposeObjectTable, config.IncludePropertyNames);

            switch (pushType)
            {
                case PushType.CreateNonExisting:
                    {
                        if (workbook.Worksheets.All(x => x.Name != sheetName))
                            success &= Create(workbook, sheetName, data, config);
                        break;
                    }
                case PushType.DeleteThenCreate:
                    {
                        if (workbook.Worksheets.Any(x => x.Name == sheetName))
                            success &= Delete(workbook, sheetName);
                        success &= Create(workbook, sheetName, data, config);
                        break;
                    }
                case PushType.UpdateOnly:
                    {
                        success &= Update(workbook, sheetName, data, config);
                        break;
                    }
                case PushType.UpdateOrCreateOnly:
                    {
                        if (workbook.Worksheets.All(x => x.Name != sheetName))
                            success &= Create(workbook, sheetName, data, config);
                        else
                            success &= Update(workbook, sheetName, data, config);
                        break;
                    }
                default:
                    {
                        BH.Engine.Base.Compute.RecordError($"Currently Excel adapter does not supports {nameof(PushType)} equal to {pushType}");
                        return new List<object>();
                    }
            }

            // Try to update the workbook properties and then save it.
            try
            {
                Update(workbook, config.WorkbookProperties);
                workbook.SaveAs(fileName);
                return success ? objects.ToList() : new List<object>();
            }
            catch (Exception e)
            {
                BH.Engine.Base.Compute.RecordError($"Finalisation and saving of the workbook failed with the following error: {e.Message}");
                return new List<object>();
            }
        }

        /***************************************************/
        /**** Private Methods                           ****/
        /***************************************************/

        private List<TableRow> ToTableRows(List<object> objects, List<string> properties, List<string> propertiesToIgnore, bool goDeep = false, bool transposeTable = false, bool showPropertyNames = true)
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

        /*******************************************/

        private static List<Dictionary<string, object>> GetPropertyDictionaries(List<object> objs, bool goDeep = false)
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



