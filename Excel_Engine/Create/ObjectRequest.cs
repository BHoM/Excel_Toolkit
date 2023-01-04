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

using BH.oM.Adapters.Excel;
using BH.oM.Base;
using BH.oM.Base.Attributes;
using System;
using System.ComponentModel;

namespace BH.Engine.Excel
{
    public static partial class Create
    {
        /*******************************************/
        /**** Public Methods                    ****/
        /*******************************************/

        [Description("Creates an ObjectRequest based on the worksheet name and range in an Excel-readable string format. The result will be of the type provided as input.")]
        [InputFromProperty("worksheet")]
        [Input("range", "Cell range in an Excel-readable string format. If not provided, collect the whole sheet.")]
        [Input("objectType", "Type of object to create from the table. If not proided, the objects will be CustomObjects.")]
        [Output("request", "CellValuesRequest created based on the input strings.")]
        public static ObjectRequest ObjectRequest(string worksheet = "", string range = "", Type objectType = null)
        {
            CellRange cellRange = null;
            if (!string.IsNullOrWhiteSpace(range))
            {
                cellRange = Create.CellRange(range);
                if (cellRange == null)
                    return null;
            }

            if (objectType == null)
                objectType = typeof(CustomObject);

            return new ObjectRequest { Worksheet = worksheet, Range = cellRange, ObjectType = objectType };
        }

        /*******************************************/
    }
}


