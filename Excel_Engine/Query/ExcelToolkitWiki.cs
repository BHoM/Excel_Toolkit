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

using BH.Engine.Reflection;
using BH.oM.Adapters.Excel;
using BH.oM.Base.Attributes;
using System;
using System.ComponentModel;
using System.Text.RegularExpressions;

namespace BH.Engine.Excel
{
    public static partial class Query
    {
        /***************************************************/
        /****              Public methods               ****/
        /***************************************************/

        [Description("Generates the Excel Toolkit wiki URL for a provided page.")]
        [Input("page", "An optional page in the wiki to link to. If no page is provided, the root URL is returned.")]
        [Output("url", "Fully qualified URL for the Excel Toolkit wiki.")]
        public static string ExcelToolkitWiki(string page = null)
        {
            string url = "https://github.com/BHoM/Excel_Toolkit/wiki";

            if (!string.IsNullOrEmpty(page))
            {
                url += $"/{page}";
            }

            return url;
        }

        /***************************************************/
    }
}




