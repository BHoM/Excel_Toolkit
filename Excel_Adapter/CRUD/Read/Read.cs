/*
 * This file is part of the Buildings and Habitats object Model (BHoM)
 * Copyright (c) 2015 - 2025, the respective contributors. All rights reserved.
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
using BH.Engine.Base;
using BH.Engine.Excel;
using BH.oM.Adapter;
using BH.oM.Adapters.Excel;
using BH.oM.Base;
using BH.oM.Data.Requests;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace BH.Adapter.Excel
{
    public partial class ExcelAdapter
    {
        /***************************************************/
        /**** Method Overrides                          ****/
        /***************************************************/

        protected override IEnumerable<IBHoMObject> Read(IRequest request, ActionConfig actionConfig = null)
        {
            XLWorkbook workbook = null;
            try
            {
                if (m_FileSettings != null)
                {
                    FileStream fileStream = new FileStream(m_FileSettings.GetFullFileName(), FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
                    workbook = new XLWorkbook(fileStream);
                    fileStream.Close();
                }
                else
                    workbook = new XLWorkbook(m_InputStream);
            }
            catch
            {
                // No error raised here because it will get raised under if (workbook == null) below.
            }

            if (workbook == null)
            {
                BH.Engine.Base.Compute.RecordError("The file under location specified in the settings is not a valid Excel workbook.");
                return new List<IBHoMObject>();
            }

            return ReadExcel(workbook, request, actionConfig);
        }


        /***************************************************/
        /**** Private Methods                           ****/
        /***************************************************/

        private List<IBHoMObject> ReadExcel(XLWorkbook workbook, IRequest request, ActionConfig actionConfig)
        {
            if (request is ObjectRequest)
            {
                List<TableRow> result = ReadExcel(workbook, ((ObjectRequest)request).Worksheet, ((ObjectRequest)request).Range, true).OfType<TableRow>().ToList();
                return CreateObjects(result, ((ObjectRequest)request).ObjectType);
            }
            else if (request is CellValuesRequest)
            {
                return ReadExcel(workbook, ((CellValuesRequest)request).Worksheet, ((CellValuesRequest)request).Range, true);
            }
            else if (request is CellContentsRequest)
                return ReadExcel(workbook, ((CellContentsRequest)request).Worksheet, ((CellContentsRequest)request).Range, false);
            else if (request is WorksheetsRequest)
                return ReadExcel(workbook, ((WorksheetsRequest)request));
            else if (request is BatchRequest batchRequest)
                return batchRequest.Requests.SelectMany(x => ReadExcel(workbook, x, actionConfig)).ToList();
            else
            {
                BH.Engine.Base.Compute.RecordError($"Requests of type {request?.GetType()} are not supported by the Excel adapter.");
                return new List<IBHoMObject>();
            }
        }

        /***************************************************/

        private List<IBHoMObject> ReadExcel(XLWorkbook workbook, WorksheetsRequest request)
        {
            IEnumerable<IXLWorksheet> worksheets = Worksheets(workbook, null);

            List<BH.oM.Adapters.Excel.Worksheet> sheets = worksheets.Select(x =>
                {
                    Worksheet sheet = new BH.oM.Adapters.Excel.Worksheet();
                    sheet.Name = x.Name;
                    return sheet;
                }).ToList();

            if (!string.IsNullOrEmpty(request.NameContains))
                sheets = sheets.Where(x => x.Name.ToLower().Contains(request.NameContains.ToLower())).ToList();

            return sheets.ToList<IBHoMObject>();
        }

        /***************************************************/

        private List<IBHoMObject> ReadExcel(XLWorkbook workbook, string worksheet, CellRange range, bool valuesOnly)
        {
            string rangeString = "";
            if (range != null)
            {
                if (string.IsNullOrEmpty(range.From.Column))
                    range.From.Column = "A";

                if (range.From.Row == -1)
                    range.From.Row = 1;

                if (string.IsNullOrEmpty(range.To.Column))
                    range.To.Column = MaximumColumnName(workbook, worksheet);

                if (range.To.Row == -1)
                    range.To.Row = MaximumRowNumber(workbook, worksheet);

                rangeString = range.ToExcel();
                if (string.IsNullOrWhiteSpace(rangeString))
                    return new List<IBHoMObject>();
            }

            return ReadExcel(workbook, worksheet, rangeString, valuesOnly);
        }

        /***************************************************/

        private List<IBHoMObject> ReadExcel(XLWorkbook workbook, string worksheet, string range, bool valuesOnly)
        {
            List<IBHoMObject> result = new List<IBHoMObject>();
            IXLWorksheet ixlWorksheet = Worksheets(workbook, worksheet).FirstOrDefault();
            if (ixlWorksheet == null)
            {
                Engine.Base.Compute.RecordError("worksheet provided cannot be found.");
                return new List<IBHoMObject>();
            }

            IXLRange ixlRange = Range(ixlWorksheet, range);
            if (ixlRange == null)
            {
                Engine.Base.Compute.RecordError("Range provided is not in the correct format for an Excel spreadsheet.");
                return new List<IBHoMObject>();
            }

            List<List<object>> table = new List<List<object>>();

            foreach (IXLRangeRow row in ixlRange.Rows())
            {
                List<object> dataRow = new List<object>();
                foreach (IXLRangeColumn column in ixlRange.Columns())
                {
                    if (valuesOnly)
                        dataRow.Add(ixlWorksheet.Cell(row.RowNumber(), column.ColumnNumber()).CellValueOrCachedValue());
                    else
                        dataRow.Add((ixlWorksheet.Cell(row.RowNumber(), column.ColumnNumber())).FromExcel());
                }

                table.Add(dataRow);
            }

            return table.Select(x => new TableRow { Content = x }).ToList<IBHoMObject>();
        }

        /***************************************************/

        private IEnumerable<IXLWorksheet> Worksheets(IXLWorkbook workbook, string worksheet)
        {
            if (!string.IsNullOrWhiteSpace(worksheet))
            {
                try
                {
                    return new List<IXLWorksheet> { workbook.Worksheet(worksheet) };
                }
                catch
                {
                    BH.Engine.Base.Compute.RecordError("No worksheets matching the request have been found.");
                    return new List<IXLWorksheet>();
                }
            }
            else
                return workbook.Worksheets;
        }

        /***************************************************/

        private IXLRange Range(IXLWorksheet worksheet, string range)
        {
            if (!string.IsNullOrWhiteSpace(range))
            {
                try
                {
                    return worksheet.Range(range);
                }
                catch
                {
                    return null;
                }
            }
            else
                return worksheet.Range(worksheet.FirstCellUsed().Address, worksheet.LastCellUsed().Address);
        }

        /***************************************************/

        private List<IBHoMObject> CreateObjects(List<TableRow> rows, Type type)
        {
            if (rows.Count < 2)
                return new List<IBHoMObject>();

            if (type == null || type == typeof(CustomObject))
                return CreateCustomObjects(rows);
            else if (!typeof(IBHoMObject).IsAssignableFrom(type))
            {
                Engine.Base.Compute.RecordError($"The type {type} is not an IBHoMObject");
                return new List<IBHoMObject>();
            }

            List<string> properties = rows.First().Content.Select(x => x.ToString()).ToList();

            return rows.Skip(1).Select(row =>
            {
                object instance = Activator.CreateInstance(type);
                for (int i = 0; i < Math.Min((int)properties.Count(), (int)row.Content?.Count()); i++)
                    instance.SetPropertyValue(properties[i], row.Content[i]);
                return instance;
            }).OfType<IBHoMObject>().ToList();
        }

        /***************************************************/

        private List<IBHoMObject> CreateCustomObjects(List<TableRow> rows)
        {
            if (rows.Count < 2)
                return new List<IBHoMObject>();

            List<string> customProperties = typeof(CustomObject).GetProperties().Select(x => x.Name).ToList();
            List<string> keys = rows.First().Content.Select(x => x.ToString()).ToList();

            return rows.Skip(1).Select(row =>
            {
                CustomObject result = new CustomObject();

                Dictionary<string, object> item = new Dictionary<string, object>();
                for (int i = 0; i < Math.Min((int)keys.Count(), (int)row.Content?.Count()); i++)
                {
                    if (customProperties.Contains(keys[i]))
                        result.SetPropertyValue(keys[i], row.Content[i]);
                    else
                        item[keys[i]] = row.Content[i];
                }

                result.CustomData = item;
                return result;

            }).ToList<IBHoMObject>();
        }

        /***************************************************/

        private string MaximumColumnName(IXLWorkbook workbook, string worksheet)
        {
            IXLWorksheet sheet = Worksheets(workbook, worksheet).FirstOrDefault();
            if (sheet == null)
                return "XFD"; //Maximum Excel Column name

            int columnNumber = sheet.LastColumnUsed().ColumnNumber();
            return ConvertToColumnName(columnNumber);
        }

        /***************************************************/

        private int MaximumRowNumber(IXLWorkbook workbook, string worksheet)
        {
            IXLWorksheet sheet = Worksheets(workbook, worksheet).FirstOrDefault();
            if (sheet == null)
                return 1048576; //Maximum Excel Row number

            return sheet.LastRowUsed().RowNumber();
        }

        /***************************************************/

        private string ConvertToColumnName(int number)
        {
            //Taken from https://frasergreenroyd.com/convert-a-number-to-an-excel-column-heading/
            int mod = 0;
            string columnHeading = "";

            while (number > 0)
            {
                mod = (number - 1) % 26;
                columnHeading = System.Convert.ToChar(65 + mod).ToString() + columnHeading;
                number = (int)((number - mod) / 26);
            }

            return columnHeading;
        }
    }
}





