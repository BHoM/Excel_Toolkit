/*
 * This file is part of the Buildings and Habitats object Model (BHoM)
 * Copyright (c) 2015 - 2024, the respective contributors. All rights reserved.
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

using BH.oM.Adapters.Excel;
using BH.oM.Base.Attributes;
using ClosedXML.Excel;
using System;
using System.ComponentModel;

namespace BH.Adapter.Excel
{
    public static partial class Create
    {
        /*******************************************/
        /**** Public Methods                    ****/
        /*******************************************/

        [Description("Converts the given ClosedXML cell contents object to a BHoM CellContents.")]
        [Input("xLCell", "ClosedXML cell contents object to convert from.")]
        [Output("cellContents", "BHoM CellContents based on the input ClosedXML cell contents object.")]
        public static CellContents FromExcel(this IXLCell xLCell)
        {
            if (xLCell == null)
                return null;

            return new CellContents()
            {
                Comment = xLCell.HasComment ? xLCell.GetComment().Text : "",
                Value = xLCell.CellValueOrCachedValue(),
                Address = BH.Engine.Excel.Create.CellAddress(xLCell.Address.ToString()),
                DataType = xLCell.DataType.SystemType(),
                FormulaA1 = xLCell.FormulaA1,
                FormulaR1C1 = xLCell.FormulaR1C1,
                HyperLink = xLCell.HasHyperlink ? xLCell.GetHyperlink().ExternalAddress.ToString() : "",
                RichText = xLCell.HasRichText ? xLCell.GetRichText().Text : ""
            };


        }

        /*******************************************/

        [Description("Gets the value of the cell, or cached value if the TryGetValue method fails. Raises a warning if the cached value is used, and ClosedXML beleives the cell needs to be recalculated.")]
        [Input("xLCell", "IXLCell to get the (cached) value from.")]
        [Input("value", "Value or cached value of the cell.")]
        public static object CellValueOrCachedValue(this IXLCell xLCell)
        {
            XLCellValue value;
            if (!xLCell.TryGetValue(out value))
            {
                //If not able to just get the value, then get the cached value
                //If cell is flagged as needing recalculation, raise warning.
                if (xLCell.NeedsRecalculation)
                    BH.Engine.Base.Compute.RecordWarning($"Cell {xLCell?.Address?.ToString() ?? "unknown"} is flagged as needing to be recalculated, but this is not able to be done. The cached value for this cell is returned, which for most cases is correct, but please check the validity of the value.");

                value = xLCell.CachedValue;
            }
            return ExtractValue(value);
        }

        /*******************************************/
        /**** Private Methods                   ****/
        /*******************************************/

        private static Type SystemType(this XLDataType dataType)
        {
            switch (dataType)
            {
                case XLDataType.Boolean:
                    return typeof(bool);
                case XLDataType.DateTime:
                    return typeof(DateTime);
                case XLDataType.Number:
                    return typeof(double);
                case XLDataType.Text:
                    return typeof(string);
                case XLDataType.TimeSpan:
                    return typeof(TimeSpan);
                default:
                    return null;
            }
        }

        /*******************************************/

        private static object ExtractValue(XLCellValue xCellValue)
        {
            switch (xCellValue.Type)
            {
                case XLDataType.Boolean:
                    return xCellValue.GetBoolean();
                case XLDataType.DateTime:
                    return xCellValue.GetDateTime();
                case XLDataType.Number:
                    return xCellValue.GetNumber();
                case XLDataType.Text:
                    return xCellValue.GetText();
                case XLDataType.TimeSpan:
                    return xCellValue.GetTimeSpan();
                case XLDataType.Error:
                    return null;
                default:
                    return null;
            }
        }

        /*******************************************/
    }
}



