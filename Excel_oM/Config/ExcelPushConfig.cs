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

using BH.oM.Adapter;
using System.Collections.Generic;
using System.ComponentModel;

namespace BH.oM.Adapters.Excel
{
    [Description("Configuration used for adapter interaction with Excel on Push action.")]
    public class ExcelPushConfig : ActionConfig
    {
        /***************************************************/
        /****             Public Properties             ****/
        /***************************************************/

        [Description("Name of the worksheet to write to.")]
        public virtual string Worksheet { get; set; } = "";

        [Description("The first cell that will be filled with the pushed objects, i.e. top-left cell of the populated space in the spreadsheet.")]
        public virtual CellAddress StartingCell { get; set; } = new CellAddress();

        [Description("List of object properties to push to the table. Those will form the columns of the created table.")]
        public virtual List<string> ObjectProperties { get; set; } = new List<string>();

        [Description("List of object properties that shouldn't be added to the table.")]
        public virtual List<string> PropertiesToIgnore { get; set; } = new List<string> { "CustomData", "Tags", "BHoM_Guid", "Fragments" };

        [Description("If true, non-primitive properties will be divided in multiple column.")]
        public virtual bool GoDeepInProperties { get; set; } = false;

        [Description("If true, transpose the table so that each column is a separate object and each row is a different property.")]
        public virtual bool TransposeObjectTable { get; set; } = false;

        [Description("If true, the property names will be shown as the first row (or column if transposed)")]
        public virtual bool IncludePropertyNames { get; set; } = true;

        [Description("Properties to apply to workbook and contents. If not null, the meta information of the workbook will be updated on push.")]
        public virtual WorkbookProperties WorkbookProperties { get; set; } = null;

        /***************************************************/
    }
}




