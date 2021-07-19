using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Mechanical;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.Style;
using Floor = Autodesk.Revit.DB.Floor;

namespace ExcelExporterImporter.Common
{
    public static class RevitUtilities
    {
        private const ScheduleFieldType CountHost = (ScheduleFieldType) 23; //<Add type 23 which is a host account

        /// <summary>
        /// </summary>
        /// <param name="e"></param>
        /// <returns></returns>
        private static double GetMaterialValumeOfElement(Element e)
        {
            var materials = e.GetMaterialIds(false);

            return materials.Sum(materialId => e.GetMaterialVolume(materialId));
        }

        /// <summary>
        /// </summary>
        /// <param name="e"></param>
        /// <returns></returns>
        private static double GetMaterialAreaOfElement(Element e)
        {
            var materials = e.GetMaterialIds(false);

            return materials.Sum(materialId => e.GetMaterialArea(materialId, false));
        }
#if REVIT2021
            /// <summary>
            /// 
            /// </summary>
            /// <param name="doc"></param>
            /// <param name="unitType"></param>
            /// <param name="value"></param>
            /// <returns></returns>
            private static double ConvertToDisplayUnit(Document doc, ForgeTypeId unitType, double value)
            {
                var fo = doc.GetUnits().GetFormatOptions(unitType);

                var dut = fo.GetUnitTypeId();

                return UnitUtils.ConvertFromInternalUnits(value, dut);
            }
#else
        /// <summary>
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="unitType"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        private static double ConvertToDisplayUnit(Document doc, UnitType unitType, double value)
        {
            var fo = doc.GetUnits().GetFormatOptions(unitType);

            var dut = fo.DisplayUnits;

            return UnitUtils.ConvertFromInternalUnits(value, dut);
        }
#endif
#if REVIT2021
            /// <summary>
            /// Get display format
            /// </summary>
            /// <param name="doc"></param>
            /// <param name="ForgeTypeId"></param>
            /// <returns>Returns the display format</returns>
            public static string GetUnitTypeSymbol(Document doc, ForgeTypeId unitType)
            {
                if (unitType.Empty())
                {
                    return "";
                }
                var fo = doc.GetUnits().GetFormatOptions(unitType);
                
                if (fo.GetSymbolTypeId().Empty())
                {
                    return "";
                }
                var formatValueOptions = new FormatValueOptions();
                formatValueOptions.SetFormatOptions(fo);
                string sResult;
                if(fo.GetUnitTypeId() == UnitTypeId.Celsius && fo.GetSymbolTypeId() == SymbolTypeId.DegreeC)//Problem with format for degrees, replace error with the correct format
                {
                    sResult = "0.00 °C";
                }
                else
                {
                    sResult = UnitFormatUtils.Format(doc.GetUnits(), unitType, 0, true, formatValueOptions);
                }
                return sResult;
            }
#else
        /// <summary>
        ///     Get display format
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="unitType"></param>
        /// <returns>Returns the display format</returns>
        public static string GetUnitTypeSymbol(Document doc, UnitType unitType)
        {
            if (unitType == UnitType.UT_Undefined) return "";
            var fo = doc.GetUnits().GetFormatOptions(unitType);
            if (fo.UnitSymbol == UnitSymbolType.UST_NONE) return "";
            var formatValueOptions = new FormatValueOptions();
            formatValueOptions.SetFormatOptions(fo);
            string sResult;
            if (fo.DisplayUnits == DisplayUnitType.DUT_CELSIUS && fo.UnitSymbol == UnitSymbolType.UST_DEGREE_C
            ) //Problem with format for degrees, replace error with the correct format
                sResult = "0.00 °C";
            else
                sResult = UnitFormatUtils.Format(doc.GetUnits(), unitType, 0, true, false, formatValueOptions);
            return sResult;
        }
#endif
        /// <summary>
        ///     Method that will get the parameters related to the cell at the symbol family level
        /// </summary>
        /// <param name="element"></param>
        /// <param name="parameterId"></param>
        /// <param name="parameter"></param>
        /// <returns>Parameter</returns>
        private static Parameter GetFamilySymbol(Element element, ElementId parameterId, Parameter parameter)
        {
            Parameter parameter2 = null;
            var FIelement = element as FamilyInstance;
            if (FIelement != null)
            {
                var familysymbol = FIelement.Symbol;
                if (familysymbol != null)
                {
                    parameter2 = familysymbol.get_Parameter((BuiltInParameter) parameterId.IntegerValue);
                    if (parameter2 != null) parameter = parameter2;
                }
            }

            return parameter;
        }

        /// <summary>
        ///     Method that will get the parameters related to the cell at the level of the wall type
        /// </summary>
        /// <param name="element"></param>
        /// <param name="parameterId"></param>
        /// <param name="parameter"></param>
        /// <returns>Parameter</returns>
        private static Parameter GetWallType(Element element, ElementId parameterId, Parameter parameter)
        {
            Parameter parameter2 = null;
            var Welement = element as Wall;
            if (Welement != null)
            {
                var walltype = Welement.WallType;
                if (walltype != null)
                {
                    parameter2 = walltype.get_Parameter((BuiltInParameter) parameterId.IntegerValue);
                    if (parameter2 != null) parameter = parameter2;
                }
            }

            return parameter;
        }

        /// <summary>
        ///     Method that will get the parameters related to the cell
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="element"></param>
        /// <param name="scheduleField"></param>
        /// <param name="pElement"></param>
        /// <returns>Return parameter</returns>
        public static Parameter GetParameter(Document doc, Element element, ScheduleField scheduleField, Phase pElement)
        {
            Parameter parameter = null;
            var parameterId = scheduleField.ParameterId;
            //
            var sColName11 = scheduleField.GetName(); //Ligne pour facilité le débugage
            var vName11 = (BuiltInParameter) parameterId.IntegerValue; //Ligne pour facilité le débugage
            var sStop121 = "asdfasdf";
            //Il va chercher les paramètres du champs pour permettre l'affichage du bon format
            //----Action selon le type de champ / Action according to the type of field----
            switch (scheduleField.FieldType)
            {
                case ScheduleFieldType.Formula:
                case ScheduleFieldType.Count:
                case CountHost:
                    break;
                case ScheduleFieldType.ElementType:
                    var ListParam = element.GetOrderedParameters();
                    parameter = ListParam.Where(i => i.Id == parameterId).FirstOrDefault();
                    if (parameter == null)
                    {
                        var ElementType = doc.GetElement(element.GetTypeId());
                        if (ElementType != null)
                        {
                            parameter = ElementType.get_Parameter((BuiltInParameter) parameterId.IntegerValue);
                            if (parameter != null)
                            {
                                var tTypeParameter = parameter.Element.GetType();
                                if (tTypeParameter.Name == "FamilyInstance")
                                {
                                }
                                else if (tTypeParameter.Name == "FamilySymbol")
                                {
                                    parameter = GetFamilySymbol(element, parameterId, parameter);
                                }
                                else if (tTypeParameter.Name == "WallType")
                                {
                                    parameter = GetWallType(element, parameterId, parameter);
                                }
                            }
                        }
                    }

                    break;
                case ScheduleFieldType.Space:
                    var familyInstanceSpace = element as FamilyInstance;
                    if (familyInstanceSpace != null)
                    {
                        var space = familyInstanceSpace.Space;
                        if (space != null) parameter = space.get_Parameter((BuiltInParameter) parameterId.IntegerValue);
                    }

                    break;
                case ScheduleFieldType.Analytical:
                    var elementanalytical = element.GetAnalyticalModel();
                    if (elementanalytical != null)
                        parameter = elementanalytical.get_Parameter((BuiltInParameter) parameterId.IntegerValue);
                    break;
                case ScheduleFieldType.ProjectInfo:
                    var vProjectInfo = doc.ProjectInformation;
                    if (vProjectInfo != null)
                        parameter = vProjectInfo.get_Parameter((BuiltInParameter) parameterId.IntegerValue);
                    break;
                case ScheduleFieldType.FromRoom:
                case ScheduleFieldType.ToRoom:
                case ScheduleFieldType.Room:
                    var familyInstance = element as FamilyInstance;
                    if (familyInstance != null)
                    {
                        var room = GetFamillyRoom(familyInstance, pElement, scheduleField.FieldType);
                        if (room != null) parameter = room.get_Parameter((BuiltInParameter) parameterId.IntegerValue);
                    }
                    else
                    {
                        var sElement = element as Space;
                        if (sElement != null)
                        {
                            var sRoom = sElement.Room;
                            if (sRoom != null)
                                parameter = sRoom.get_Parameter((BuiltInParameter) parameterId.IntegerValue);
                        }
                    }

                    break;
                case ScheduleFieldType.MaterialQuantity:
                    parameter = element.get_Parameter((BuiltInParameter) parameterId.IntegerValue);
                    break;
                case ScheduleFieldType.PhysicalInstance:
                    parameter = element.get_Parameter((BuiltInParameter) parameterId.IntegerValue);
                    break;
                case ScheduleFieldType.StructuralMaterial:
                    var SM_familyInstance = element as FamilyInstance;
                    var SM_Wall = element as Wall;
                    var SM_FloorElement = element as Floor;

                    if (SM_familyInstance != null)
                    {
                        var Material_FamilyInstance =
                            doc.GetElement(SM_familyInstance.StructuralMaterialId) as Material;
                        if (Material_FamilyInstance != null)
                        {
                            var SAI_Material_FamilyInstance = doc.GetElement(Material_FamilyInstance.StructuralAssetId);
                            if (SAI_Material_FamilyInstance != null)
                                parameter = SAI_Material_FamilyInstance.get_Parameter(
                                    (BuiltInParameter) parameterId.IntegerValue);
                            if (parameter == null)
                                parameter = Material_FamilyInstance.get_Parameter(
                                    (BuiltInParameter) parameterId.IntegerValue);
                        }
                        else
                        {
                            var SM_FamilySymbol = SM_familyInstance.Symbol;
                            if (SM_FamilySymbol != null && parameter == null)
                            {
                                //On va chercher le Element Id dans Maétriau Structurel
                                var ParameterFamilySymbol =
                                    SM_FamilySymbol.get_Parameter(BuiltInParameter.STRUCTURAL_MATERIAL_PARAM);
                                if (ParameterFamilySymbol != null)
                                {
                                    var SM_Material_FamilySymbol =
                                        doc.GetElement(ParameterFamilySymbol.AsElementId()) as Material;
                                    if (SM_Material_FamilySymbol != null)
                                    {
                                        var SAI_Material_FamilySymbol =
                                            doc.GetElement(SM_Material_FamilySymbol.StructuralAssetId);
                                        if (SAI_Material_FamilySymbol != null)
                                            parameter = SAI_Material_FamilySymbol.get_Parameter(
                                                (BuiltInParameter) parameterId.IntegerValue);
                                        if (parameter == null)
                                            parameter = SM_Material_FamilySymbol.get_Parameter(
                                                (BuiltInParameter) parameterId.IntegerValue);
                                    }
                                }
                            }
                        }
                    }

                    if (SM_FloorElement != null && parameter == null)
                    {
                        var Material_FloorType =
                            doc.GetElement(SM_FloorElement.FloorType.StructuralMaterialId) as Material;
                        if (Material_FloorType != null)
                        {
                            var SAI_Material_FloorType = doc.GetElement(Material_FloorType.StructuralAssetId);
                            if (SAI_Material_FloorType != null)
                                parameter = SAI_Material_FloorType.get_Parameter(
                                    (BuiltInParameter) parameterId.IntegerValue);
                            if (parameter == null)
                                parameter = Material_FloorType.get_Parameter(
                                    (BuiltInParameter) parameterId.IntegerValue);
                        }
                    }

                    if (SM_Wall != null && parameter == null)
                    {
                        var SM_WallType = doc.GetElement(SM_Wall.GetTypeId());
                        if (SM_WallType != null)
                        {
                            var ParameterWallType =
                                SM_WallType.get_Parameter(BuiltInParameter.STRUCTURAL_MATERIAL_PARAM);
                            if (ParameterWallType != null)
                            {
                                var Material_WallType = doc.GetElement(ParameterWallType.AsElementId()) as Material;
                                if (Material_WallType != null)
                                {
                                    var SAI_Material_WallType = doc.GetElement(Material_WallType.StructuralAssetId);
                                    if (SAI_Material_WallType != null)
                                        parameter = SAI_Material_WallType.get_Parameter(
                                            (BuiltInParameter) parameterId.IntegerValue);
                                    if (parameter == null)
                                        parameter = Material_WallType.get_Parameter(
                                            (BuiltInParameter) parameterId.IntegerValue);
                                }
                            }
                        }
                    }
#if REVIT2016 || REVIT2017 || REVIT2018 || REVIT2019 || REVIT2020 || REVIT2021
                    var SM_WallFoundation = element as WallFoundation;
                    if(SM_WallFoundation != null && parameter == null)
                    {
                        Element SM_WallFoundationType = doc.GetElement(SM_WallFoundation.GetTypeId());
                        if (SM_WallFoundationType != null)
                        {
                            Autodesk.Revit.DB.Parameter ParameterWallFoundationType =
 SM_WallFoundationType.get_Parameter(BuiltInParameter.STRUCTURAL_MATERIAL_PARAM);
                            if(ParameterWallFoundationType != null)
                            {
                                Material Material_WallFoundationType =
 doc.GetElement(ParameterWallFoundationType.AsElementId()) as Material;
                                if(Material_WallFoundationType != null)
                                {
                                    Element SAI_Material_WallFoundationType =
 doc.GetElement(Material_WallFoundationType.StructuralAssetId);
                                    if(SAI_Material_WallFoundationType != null)
                                    {
                                        parameter =
 SAI_Material_WallFoundationType.get_Parameter((BuiltInParameter)parameterId.IntegerValue);
                                    }
                                    if(parameter == null)
                                    {
                                        parameter =
 Material_WallFoundationType.get_Parameter((BuiltInParameter)parameterId.IntegerValue);
                                    }
                                }
                            }
                        }
                    }
#endif
                    break;
                case ScheduleFieldType.Instance: //Éventuellement simplifié ce code
                    if (parameterId.IntegerValue > 0) //Indique que c'est un paramètre partagé
                    {
                        var ElementType = doc.GetElement(element.GetTypeId());
                        if (ElementType != null)
                        {
                            parameter = ElementType.get_Parameter((BuiltInParameter) parameterId.IntegerValue);
                            if (parameter != null)
                            {
                                var tTypeParameter = parameter.Element.GetType();
                                if (tTypeParameter.Name == "FamilyInstance")
                                {
                                }
                                else if (tTypeParameter.Name == "FamilySymbol")
                                {
                                    parameter = GetFamilySymbol(element, parameterId, parameter);
                                }
                                else if (tTypeParameter.Name == "WallType")
                                {
                                    parameter = GetWallType(element, parameterId, parameter);
                                }
                            }
                        }

                        if (parameter == null || !parameter.HasValue)
                        {
                            parameter = element.get_Parameter((BuiltInParameter) parameterId.IntegerValue);

                            if (parameter != null)
                            {
                                var tTypeParameter = parameter.Element.GetType();
                                if (tTypeParameter.Name == "FamilyInstance")
                                {
                                    if (parameter.IsShared && parameter.IsReadOnly)
                                        parameter = GetFamilySymbol(element, parameterId, parameter);
                                }
                                else if (tTypeParameter.Name == "FamilySymbol")
                                {
                                    parameter = GetFamilySymbol(element, parameterId, parameter);
                                }
                            }
                        }
                    }
                    else //Ce n'est pas un paramètre partagé
                    {
                        parameter = element.get_Parameter((BuiltInParameter) parameterId.IntegerValue);
                        if (!ValidValue(parameter))
                        {
                            var AM_Element = element.GetAnalyticalModel();
                            if (AM_Element != null)
                                parameter = AM_Element.get_Parameter((BuiltInParameter) parameterId.IntegerValue);
                            if (parameter != null)
                            {
                                var tTypeParameter = parameter.Element.GetType();
                                if (tTypeParameter.Name == "FamilyInstance")
                                {
                                    if (parameter.IsShared && parameter.IsReadOnly)
                                        parameter = GetFamilySymbol(element, parameterId, parameter);
                                }
                                else if (tTypeParameter.Name == "FamilySymbol")
                                {
                                    parameter = GetFamilySymbol(element, parameterId, parameter);
                                }
                            }
                        }
                    }

                    break;
                default:
                    try
                    {
                        var sColName = scheduleField.GetName(); //Ligne pour facilité le débugage
                        var vName = (BuiltInParameter) parameterId.IntegerValue; //Ligne pour facilité le débugage
                        var AM_Element = element.GetAnalyticalModel();
                        if (AM_Element != null)
                            parameter = AM_Element.get_Parameter((BuiltInParameter) parameterId.IntegerValue);
                        if (parameter == null)
                            parameter = element.get_Parameter((BuiltInParameter) parameterId.IntegerValue);
                        if (parameter != null) //Pour obtenir les valeurs des paramètres partagé
                        {
                            var tTypeParameter = parameter.Element.GetType();
                            if (tTypeParameter.Name == "FamilyInstance")
                            {
                                if (parameter.Definition.ParameterGroup == BuiltInParameterGroup.INVALID
                                ) //Clause a des fin de tests seulement
                                {
                                }

                                if (parameter.IsShared && parameter.IsReadOnly)
                                    parameter = GetFamilySymbol(element, parameterId, parameter);
                            }
                            else if (tTypeParameter.Name == "FamilySymbol")
                            {
                                parameter = GetFamilySymbol(element, parameterId, parameter);
                            }
                        }
                        else
                        {
                            //Valide si c'est un paramètre de type
                            var ElementType2 = doc.GetElement(element.GetTypeId());
                            if (ElementType2 != null)
                            {
                                parameter = ElementType2.get_Parameter((BuiltInParameter) parameterId.IntegerValue);
                                if (parameter != null)
                                {
                                    var tTypeParameter = parameter.Element.GetType();
                                    if (tTypeParameter.Name == "FamilyInstance")
                                    {
                                    }
                                    else if (tTypeParameter.Name == "FamilySymbol")
                                    {
                                        parameter = GetFamilySymbol(element, parameterId, parameter);
                                    }
                                    else if (tTypeParameter.Name == "WallType")
                                    {
                                        parameter = GetWallType(element, parameterId, parameter);
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception)
                    {
                    }

                    break;
            }

            //----End of the action according to the type of field----
            return parameter;
        }

        /// <summary>
        ///     indicates whether the parameter variable contains a value
        /// </summary>
        /// <param name="parameter"></param>
        /// <returns></returns>
        public static bool ValidValue(Parameter parameter)
        {
            var bResult = false;
            if (parameter != null)
                if (parameter.HasValue)
                    bResult = true;
            return bResult;
        }

        /// <summary>
        ///     Method for obtaining cell rights
        /// </summary>
        /// <param name="parameter"></param>
        /// <param name="scheduleField"></param>
        /// <param name="readonlyParameters"></param>
        /// <returns>True or False</returns>
        public static bool GetIsReadOnly(Parameter parameter, ScheduleField scheduleField, Hashtable readonlyParameters)
        {
            var readonlyParameter = true;
            var bExtSchedule = false; // Indique que l'information provient d'une nomenclature parent
            //----Action selon le type de champ / Action according to the type of field----
            switch (scheduleField.FieldType)
            {
                case ScheduleFieldType.Formula:
                case ScheduleFieldType.Count:
                case CountHost:
                case ScheduleFieldType.Space:
                case ScheduleFieldType.Analytical:
                case ScheduleFieldType.ProjectInfo:
                case ScheduleFieldType.FromRoom:
                case ScheduleFieldType.ToRoom:
                case ScheduleFieldType.Room:
                case ScheduleFieldType.MaterialQuantity:
                case ScheduleFieldType.StructuralMaterial:
                    bExtSchedule = true;
                    break;
            }

            //----End of the action according to the type of field----
            if (parameter == null)
            {
                readonlyParameter = true;
            }
            else
            {
                if (readonlyParameters.ContainsKey(scheduleField.ParameterId.IntegerValue) || parameter.IsReadOnly ||
                    bExtSchedule)
                    readonlyParameter = true;
                else
                    readonlyParameter = false;
            }

            return readonlyParameter;
        }

        /// <summary>
        ///     Gets the value for a schedule parameter
        /// </summary>
        /// <param name="parameter"></param>
        /// <param name="scheduleField"></param>
        /// <param name="doc"></param>
        /// <param name="element"></param>
        /// <param name="appliedParameters"></param>
        /// <returns>Returns the value or null value</returns>
        public static object GetParameterValue(Parameter parameter, ScheduleField scheduleField, Document doc,
            Element element, List<ParameterTranslation> appliedParameters)
        {
            object cellVal = null;
            var parameterId = scheduleField.ParameterId;
            switch (scheduleField.FieldType)
            {
                case ScheduleFieldType.Formula:
                    cellVal = "0";
                    break;
                case ScheduleFieldType.Count:
                case CountHost:
                    cellVal = "1";
                    break;
                case ScheduleFieldType.MaterialQuantity:
                    if (parameterId.IntegerValue == (int) BuiltInParameter.MATERIAL_AREA)
                    {
#if REVIT2021
                                cellVal =
 RevitUtilities.ConvertToDisplayUnit(doc, scheduleField.GetSpecTypeId(), RevitUtilities.GetMaterialAreaOfElement(element));
#else
                        cellVal = ConvertToDisplayUnit(doc, scheduleField.UnitType, GetMaterialAreaOfElement(element));
#endif
                    }
                    else if (parameterId.IntegerValue == (int) BuiltInParameter.MATERIAL_VOLUME)
                    {
#if REVIT2021
                                cellVal =
 RevitUtilities.ConvertToDisplayUnit(doc, scheduleField.GetSpecTypeId(), RevitUtilities.GetMaterialValumeOfElement(element));
#else
                        cellVal = ConvertToDisplayUnit(doc, scheduleField.UnitType,
                            GetMaterialValumeOfElement(element));
#endif
                    }
                    else if (parameterId.IntegerValue == (int) BuiltInParameter.MATERIAL_ASPAINT)
                    {
                        cellVal = "No";
                        if (element.GetMaterialIds(true).Count > 0) cellVal = "Yes";
                    }
                    else if (parameterId.IntegerValue == (int) BuiltInParameter.PHY_MATERIAL_PARAM_UNIT_WEIGHT)
                    {
                        double dCellVal = 0;
                        var materials = element.GetMaterialIds(false);
                        if (materials.Count > 0)
                            foreach (var Item in materials)
                            {
                                var eMaterial = doc.GetElement(Item) as Material;
                                if (eMaterial != null)
                                {
                                    var pseProperty = doc.GetElement(eMaterial.StructuralAssetId) as PropertySetElement;
                                    if (pseProperty != null)
                                    {
                                        parameter = pseProperty.get_Parameter(
                                            (BuiltInParameter) parameterId.IntegerValue);
                                        if (parameter != null)
                                        {
                                            dCellVal += Convert.ToDouble(GetParameterValue(doc, parameter,
                                                scheduleField));
                                            parameter = null;
                                        }
                                    }
                                }
                            }

                        cellVal = dCellVal;
                    }

                    break;
            }

            //---- End of the action according to the type of field----
            if (parameter != null)
            {
                cellVal = GetParameterValue(doc, parameter, scheduleField);
                // Translate parameter value based on XML settings
                var appliedTranslations = appliedParameters.Where(p => p.Name == parameter.Definition.Name).ToList();
                cellVal = appliedTranslations.Any()
                    ? TranslateValueToText(Convert.ToString(cellVal), appliedTranslations)
                    : cellVal;
            }

            return cellVal;
        }

        /// <summary>
        ///     Gets the right value depending on its type of storage
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="parameter"></param>
        /// <param name="scheduleField"></param>
        /// <returns>Returns the good value</returns>
        public static object GetParameterValue(Document doc, Parameter parameter, ScheduleField scheduleField = null)
        {
            object val = null;
            switch (parameter.StorageType)
            {
                case StorageType.Double:
                    var dVal = parameter.AsProjectUnitTypeDouble(scheduleField);
                    try
                    {
#if REVIT2021
                            if (parameter.GetUnitTypeId() == UnitTypeId.Percentage)
                            {
                                val = dVal / 100.0;
                            }
                            else
                            {
                                val = dVal;
                            }
#else
                        if (parameter.DisplayUnitType == DisplayUnitType.DUT_PERCENTAGE)
                            val = dVal / 100.0;
                        else
                            val = dVal;
#endif
                    }
                    catch (Exception)
                    {
                        val = dVal;
                    }

                    break;
                case StorageType.String:
                    val = parameter.AsString();
                    if (val == null) //Adding an empty string to replace null solved sorting problems
                        val = string.Empty;
                    break;
                case StorageType.Integer:
                    val = parameter.AsInteger();
                    if (parameter.Definition.ParameterType == ParameterType.YesNo)
                        val = (int) val == 0 ? "False" : "True";
                    break;
                case StorageType.ElementId:
                    var elementId = parameter.AsElementId();
                    if (elementId.IntegerValue < 0)
                    {
                        var cat = doc.Settings.Categories.get_Item((BuiltInCategory) elementId.IntegerValue);
                        if (cat != null) val = cat.Name;
                    }
                    else
                    {
                        if (parameter.Id.IntegerValue == -1002051 || parameter.Id.IntegerValue == -1002052)
                        {
                            var elementType = doc.GetElement(elementId) as ElementType;
                            if (elementType != null)
                            {
                                var familyName = GetElementFamilyName(doc, elementType);
                                if (parameter.Id.IntegerValue == -1002052)
                                    val = familyName + ": " + elementType.Name;
                                else
                                    val = familyName;
                            }
                        }
                        else if (parameter.Id.IntegerValue == -1012701) //Pour un type area
                        {
                            val = parameter.AsValueString();
                        }
                        else
                        {
                            var element = doc.GetElement(elementId);
                            if (element != null) val = element.Name;
                        }
                    }

                    break;
            }

            return val;
        }

        /// <summary>
        ///     Gets the list of columns that is write-only.
        /// </summary>
        /// <param name="parametersSettings"></param>
        /// <returns>Returns the list</returns>
        public static Hashtable GetListReadOnlyParamater(ParametersSettings parametersSettings)
        {
            var readonlyParameters = new Hashtable();
            foreach (var parameter in parametersSettings.ReadonlyParameters) readonlyParameters[parameter.Id] = true;
            //============================================Added column to manually lock==========================================================
            /*
                -1152384 = Image type
                -1012101 = Phase demolished
                -1012100 = Phase created
                -1002064 = Top level
                -1002063 = Base level
                -1001305 = Rough Width
                -1001304 = Rough Heigh
                -1012201 = Scope Box
                -1007110 = Story above
                -1012030 = Conceptual Types
                -1012023 = Graphical Appearance
                -1007235 = Niveau supérieur multiétage
                -1007201 = Niveau supérieur
                -1007200 = Niveau de base
                -1006922 = Limite supérieure
                -1151210 = Type de support gauche
                -1151209 = Type de support droit
                -1151208 = Type de palier
                -1151207 = Type de volée
                -1008620 = Niveau de base
                -1005500 = Matériau structurel
                -1001107 = Contrainte inférieure
                -1001105 = Hauteur non contrainte
                -1001103 = Contrainte supérieure
                -1140036 = Type de fil
                -1140334 = Type de système
                -1150115 = Dépréciation due aux impuretés sur le luminaire
                -1150114 = Dépréciation de lumen de la lampe
                -1150113 = Perte de dépréciation de la surface
                -1150112 = Perte d'inclinaison de la lampe
                -1150110 = Perte de tension
                -1114251 = Type de construction
                -1114172 = Type d'espace
                -1018257 = Type de barre mineure inférieure/intérieure
                -1018256 = Type de barre majeure inférieure/intérieure
                -1018255 = Type de barre mineure supérieure/extérieure
                -1018254 = Type de barre majeure supérieure/extérieure
                -1018023 = Correspondances de couches mineures supérieures et inférieures
                -1018022 = Correspondances de couches majeures supérieures et inférieures
                -1018021 = Correspondances de couches majeures et mineures inférieures
                -1018020 = Correspondances de couches majeures et mineures supérieures
                -1012809 = Sous-catégorie de murs
                -1012800 = Profil
                -1002107 = Matériau
                -1017701 = Panneau de treillis
                -1017733 = Clé de famille partagée
                -1017604 = Type de treillis sens répartition
                -1017603 = Type de treillis sens porteur
                -1012819 = Profil
                -1012836 = Profil
                -1152335 = Niveau de base
                -1018361 = Barre principale
                -1018305 = Barre principale
                -1018300 = Face
                -1015000 = Cas de charge
                -1013411 = Type de poutre
                -1012701 = Type de surface
                -1017705 = Emplacement
                -1152385 = Physique: Image
                -1140217 = Température du fluide
            */
            int[] tAddColReadOnly =
            {
                -1152384, -1012101, -1012100, -1002064, -1002063, -1001305, -1001304, -1012201, -1007110, -1012030,
                -1012023, -1007235, -1007201, -1007200, -1006922, -1151210, -1151209, -1151208, -1151207, -1008620,
                -1005500, -1001107, -1001105, -1001103, -1140036, -1140334, -1150115, -1150114, -1150113, -1150112,
                -1150110, -1114251, -1114172, -1018257, -1018256, -1018255, -1018254, -1018023, -1018022, -1018021,
                -1018020, -1012809, -1012800, -1002107, -1017701, -1017733, -1017604, -1017603, -1012819, -1012836,
                -1152335, -1018361, -1018305, -1018300, -1015000, -1013411, -1012701, -1017705, -1152385, -1140217
            };
            foreach (var iIdCol in tAddColReadOnly) readonlyParameters[iIdCol] = true;
            return readonlyParameters;
        }

        /// <summary>
        ///     Method to change the value of a parameter
        /// </summary>
        /// <param name="parameter"></param>
        /// <param name="value"></param>
        /// <param name="scheduleField"></param>
        public static void SetParameterValue(Parameter parameter, object value, ScheduleField scheduleField = null)
        {
            switch (parameter.StorageType)
            {
                case StorageType.Double:
                    if (value == null) value = 0;
                    var dVal = double.Parse(value.ToString());
                    var dValueAct = parameter.AsProjectUnitTypeDouble(scheduleField);
                    try
                    {
#if REVIT2021
                            if (parameter.GetUnitTypeId() == UnitTypeId.Percentage)
                            {
                                dVal *= 100.0;
                            }
#else
                        if (parameter.DisplayUnitType == DisplayUnitType.DUT_PERCENTAGE) dVal *= 100.0;
#endif
                    }
                    catch (Exception)
                    {
                        // ignored
                    }

                    if (Math.Round(dValueAct, 4) != Math.Round(dVal, 4))
                        if (!parameter.Set(parameter.ToProjectUnitType(dVal, scheduleField)))
                            throw new TargetInvocationException(
                                string.Format(Resources.ErrorOccurredWhileSettingParameter, parameter.Definition.Name,
                                    value), null);
                    break;

                case StorageType.String:
                    var strVal = Convert.ToString(parameter.AsString());
                    var bInsertString = false;
                    if (strVal != null)
                    {
                        if (!strVal.Equals(Convert.ToString(value), StringComparison.InvariantCultureIgnoreCase))
                            bInsertString = true;
                    }
                    else if (value != null && value.ToString() != string.Empty)
                    {
                        bInsertString = true;
                    }

                    if (bInsertString)
                        if (!parameter.Set(Convert.ToString(value))) //Insertion du string
                            throw new TargetInvocationException(
                                string.Format(Resources.ErrorOccurredWhileSettingParameter, parameter.Definition.Name,
                                    value), null);

                    break;

                case StorageType.Integer:
                    int iVal;
                    if (parameter.Definition.ParameterType == ParameterType.YesNo)
                        iVal = Convert.ToBoolean(value) ? 1 : 0;
                    else
                        iVal = Convert.ToInt32(value);

                    if (iVal != parameter.AsInteger())
                        //if (!parameter.SetValueString(iVal.ToString()))//Insertion de la valeur
                        if (!parameter.Set(iVal)) //Insertion de la valur
                            throw new TargetInvocationException(
                                string.Format(Resources.ErrorOccurredWhileSettingParameter, parameter.Definition.Name,
                                    iVal), null);
                    break;

                case StorageType.ElementId:
                    var doc = parameter.Element.Document;
                    FilteredElementCollector collector = null;
                    Element newElement = null;

                    if (parameter.Id.IntegerValue == (int) BuiltInParameter.ELEM_TYPE_PARAM)
                    {
                        collector = new FilteredElementCollector(doc).WhereElementIsElementType();
                        newElement =
                            collector.FirstOrDefault(e => e.Name.Trim().Equals(Convert.ToString(value).Trim()));
                    }
                    else if (parameter.Id.IntegerValue == (int) BuiltInParameter.ELEM_FAMILY_PARAM)
                    {
                        collector = new FilteredElementCollector(doc).WhereElementIsElementType();
                        newElement = collector.Cast<ElementType>().FirstOrDefault(elementType =>
                            GetElementFamilyName(doc, elementType).Trim().Equals(Convert.ToString(value).Trim(),
                                StringComparison.CurrentCultureIgnoreCase));
                    }
                    else
                    {
                        return;
                    }

                    var originalElement = doc.GetElement(parameter.AsElementId());
                    if (originalElement != null && newElement == null)
                    {
                        if (!parameter.Set(new ElementId(BuiltInParameter.INVALID))) //Insertion de la valeur
                            throw new TargetInvocationException(
                                string.Format(Resources.ErrorOccurredWhileSettingParameter, parameter.Definition.Name,
                                    "(Invalid)"), null);
                    }
                    else if (newElement != null)
                    {
                        if (originalElement != null && originalElement.Id.IntegerValue != newElement.Id.IntegerValue ||
                            originalElement == null)
                            if (!parameter.Set(newElement.Id)) //Insertion de la valeur
                                throw new TargetInvocationException(
                                    string.Format(Resources.ErrorOccurredWhileSettingParameter,
                                        parameter.Definition.Name, newElement.Id), null);
                    }

                    break;
            }
        }

        /// <summary>
        ///     Method for obtaining the current phase for an element
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="element"></param>
        /// <returns>Null or Phase object</returns>
        public static Phase GetElementPhase(Document doc, Element element)
        {
            Phase pPhase = null;
            var ParameterPhase = element.get_Parameter(BuiltInParameter.PHASE_CREATED);
            if (ParameterPhase != null)
            {
                var elementId = ParameterPhase.AsElementId();
                foreach (Phase phase in doc.Phases)
                    if (phase.Id.IntegerValue == elementId.IntegerValue)
                        pPhase = phase;
            }

            return pPhase;
        }

        /// <summary>
        ///     Method thast gets Familly Room
        /// </summary>
        /// <param name="familyInstance"></param>
        /// <param name="pPhase"></param>
        /// <param name="fieldType"></param>
        /// <returns>Null or Room object</returns>
        private static Room GetFamillyRoom(FamilyInstance familyInstance, Phase pPhase, ScheduleFieldType fieldType)
        {
            Room room = null;
            if (pPhase != null)
            {
                if (fieldType == ScheduleFieldType.FromRoom)
                    room = familyInstance.get_FromRoom(pPhase);
                else if (fieldType == ScheduleFieldType.ToRoom)
                    room = familyInstance.get_ToRoom(pPhase);
                else if (fieldType == ScheduleFieldType.Room) room = familyInstance.get_Room(pPhase);
            }
            else
            {
                if (fieldType == ScheduleFieldType.FromRoom)
                {
                    //room = familyInstance.FromRoom;
                }
                else if (fieldType == ScheduleFieldType.ToRoom)
                {
                    //room = familyInstance.ToRoom;
                }
                else if (fieldType == ScheduleFieldType.Room)
                {
                    //room = familyInstance.Room;
                }
            }

            return room;
        }

        /// <summary>
        ///     Translate value to text
        /// </summary>
        /// <param name="value"></param>
        /// <param name="parameters"></param>
        /// <returns>Result translate</returns>
        public static string TranslateValueToText(string value, List<ParameterTranslation> parameters)
        {
            foreach (var parameter in parameters)
            foreach (var translation in parameter.Translations)
                if (translation.Value.ToLower().Trim() == value.ToLower().Trim())
                    return translation.Text.ToLower().Trim();
            return value;
        }

        /// <summary>
        ///     Translate text to value
        /// </summary>
        /// <param name="text"></param>
        /// <param name="parameters"></param>
        /// <returns>Result value</returns>
        public static string TranslateTextToValue(string text, List<ParameterTranslation> parameters)
        {
            foreach (var parameter in parameters)
            foreach (var translation in parameter.Translations)
                if (translation.Text.ToLower().Trim() == text.ToLower().Trim())
                    return translation.Value.ToLower().Trim();

            throw new ArgumentOutOfRangeException("text", string.Format(Resources.UnableFindCorrespondingValue, text));
        }

        /// <summary>
        ///     Valid if the parameter can be exported
        /// </summary>
        /// <param name="parameter"></param>
        /// <param name="ignoredParameters"></param>
        /// <param name="location"></param>
        /// <returns>True or False</returns>
        public static bool CanExportParameter(Parameter parameter, List<IgnoredParameters> ignoredParameters,
            string location)
        {
            if (ignoredParameters.Any(p => (p.Location == "*" || p.Location.ToLower() == location.ToLower()) &&
                                           p.Parameter.Any(pr =>
                                               pr.Name.ToLower() == parameter.Definition.Name.ToLower())))
                return false;

            if (parameter.Definition.Name.ToLower().Contains("none")) return false;

            return true;
        }

        /// <summary>
        ///     Valid if the parameter can be exported
        /// </summary>
        /// <param name="scheduleField"></param>
        /// <param name="ignoredParameters"></param>
        /// <param name="location"></param>
        /// <returns></returns>
        public static bool CanExportParameter(ScheduleField scheduleField, List<IgnoredParameters> ignoredParameters,
            string location)
        {
            if (ignoredParameters.Any(p =>
                (p.Location == "*" || p.Location.Equals(location, StringComparison.InvariantCultureIgnoreCase)) &&
                p.Parameter.Any(pr =>
                    pr.Name.Equals(scheduleField.GetName(), StringComparison.InvariantCultureIgnoreCase))))
                return false;

            return true;
        }

        /// <summary>
        ///     Convert RGB value to color object
        /// </summary>
        /// <param name="colorString"></param>
        /// <returns>Color</returns>
        public static Color ConvertRgbToColor(string colorString)
        {
            if (!string.IsNullOrEmpty(colorString))
            {
                var rgb = colorString.Split(',');
                if (rgb.Length == 3)
                    try
                    {
                        return new Color(Convert.ToByte(rgb[0]), Convert.ToByte(rgb[1]), Convert.ToByte(rgb[2]));
                    }
                    catch (FormatException)
                    {
                    }
                    catch (OverflowException)
                    {
                    }
            }

            return null;
        }

        /// <summary>
        ///     A method that inserts a row at the end of the table indicating that this table can not be imported.
        /// </summary>
        /// <param name="worksheet"></param>
        public static void InsertMsgNotBeImported(ExcelWorksheet worksheet)
        {
            var iColCnt = worksheet.Dimension.End.Column;
            var iRowCnt = worksheet.Dimension.End.Row;
            var row = iRowCnt + 1;
            worksheet.Cells[row, 1, row, iColCnt].Merge = true;
            worksheet.Cells[row, 1].Value = Resources.ThisSheetCanNotBeImported;
            worksheet.Cells[row, 1, row, iColCnt].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row, 1, row, iColCnt].Style.Fill.BackgroundColor
                .SetColor(Styles.BackgroundColor.MsgNotBeImported);
            worksheet.Cells[row, 1, row, iColCnt].Style.Font.Color.SetColor(Styles.FontColor.MsgNotBeImported);
        }

        /// <summary>
        ///     Method that locks all columns in a table
        /// </summary>
        /// <param name="worksheet"></param>
        public static void LockAllColumns(ExcelWorksheet worksheet)
        {
            for (var x = 1; x <= worksheet.Dimension.End.Column; x++) LockColumn(worksheet, x);
        }

        /// <summary>
        ///     Method that locks a column in an excel table
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="iCol"></param>
        public static void LockColumn(ExcelWorksheet worksheet, int iCol)
        {
            worksheet.Column(iCol).Style.Locked = true;
            worksheet.Column(iCol).Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Column(iCol).Style.Fill.BackgroundColor.SetColor(Styles.BackgroundColor.CellLocked);
            worksheet.Column(iCol).Style.Font.Color.SetColor(Styles.FontColor.CellLocked);
        }

        /// <summary>
        ///     Method that locks and formats an ElementType column
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="iCol"></param>
        public static void FormattingColElementType(ExcelWorksheet worksheet, int iCol)
        {
            worksheet.Column(iCol).Style.Locked = true;
            worksheet.Column(iCol).Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Column(iCol).Style.Fill.BackgroundColor.SetColor(Styles.BackgroundColor.ColElementType);
            worksheet.Column(iCol).Style.Font.Color.SetColor(Styles.FontColor.ColElementType);
        }

        /// <summary>
        ///     Modify the format of a formula column
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="iCol"></param>
        public static void FormattingColFormula(ExcelWorksheet worksheet, int iRowStart, int iRowEnd, int iCol)
        {
            worksheet.Cells[iRowStart, iCol, iRowEnd, iCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[iRowStart, iCol, iRowEnd, iCol].Style.Fill.BackgroundColor
                .SetColor(Styles.BackgroundColor.TypeFormula);
            worksheet.Cells[iRowStart, iCol, iRowEnd, iCol].Style.Font.Color.SetColor(Styles.FontColor.TypeFormula);
        }

        /// <summary>
        ///     Format the cell that contains the warning for formulas
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="iRowStart"></param>
        /// <param name="iColStart"></param>
        /// <param name="iRowEnd"></param>
        /// <param name="iColEnd"></param>
        public static void FormattingRowWarningFormula(ExcelWorksheet worksheet, int iRowStart, int iColStart,
            int iRowEnd, int iColEnd)
        {
            worksheet.Cells[iRowStart, iColStart, iRowEnd, iColEnd].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[iRowStart, iColStart, iRowEnd, iColEnd].Style.Fill.BackgroundColor
                .SetColor(Styles.BackgroundColor.TypeFormula);
            worksheet.Cells[iRowStart, iColStart, iRowEnd, iColEnd].Style.Font.Color
                .SetColor(Styles.FontColor.TypeFormula);

            worksheet.Cells[iRowStart, iColStart, iRowEnd, iColEnd].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[iRowStart, iColStart, iRowEnd, iColEnd].Style.Border.Left.Color
                .SetColor(Styles.BorderColor.TypeFormula);

            worksheet.Cells[iRowStart, iColStart, iRowEnd, iColEnd].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[iRowStart, iColStart, iRowEnd, iColEnd].Style.Border.Right.Color
                .SetColor(Styles.BorderColor.TypeFormula);

            worksheet.Cells[iRowStart, iColStart, iRowEnd, iColEnd].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[iRowStart, iColStart, iRowEnd, iColEnd].Style.Border.Top.Color
                .SetColor(Styles.BorderColor.TypeFormula);

            worksheet.Cells[iRowStart, iColStart, iRowEnd, iColEnd].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[iRowStart, iColStart, iRowEnd, iColEnd].Style.Border.Bottom.Color
                .SetColor(Styles.BorderColor.TypeFormula);
        }

        /// <summary>
        ///     Method that applies the formatting of the header
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="iRow"></param>
        public static void FormattingTheHeader(ExcelWorksheet worksheet, int iRow)
        {
            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Fill.BackgroundColor
                .SetColor(Styles.BackgroundColor.Header);
            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Font.Color
                .SetColor(Styles.FontColor.Header);

            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Border.Left.Color
                .SetColor(Styles.BorderColor.HeaderCell);

            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Border.Right.Style =
                ExcelBorderStyle.Thin;
            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Border.Right.Color
                .SetColor(Styles.BorderColor.HeaderCell);

            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Border.Top.Color
                .SetColor(Styles.BorderColor.Header);

            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Border.Bottom.Style =
                ExcelBorderStyle.Thick;
            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Border.Bottom.Color
                .SetColor(Styles.BorderColor.Header);

            worksheet.Cells[iRow, 1].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[iRow, 1].Style.Border.Left.Color.SetColor(Styles.BorderColor.Header);

            worksheet.Cells[iRow, worksheet.Dimension.Columns].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[iRow, worksheet.Dimension.Columns].Style.Border.Right.Color
                .SetColor(Styles.BorderColor.Header);
        }

        /// <summary>
        ///     Method that formats the Excel table
        /// </summary>
        /// <param name="worksheet"></param>
        public static void FormatingTable(ExcelWorksheet worksheet)
        {
            worksheet.Cells[1, 1, worksheet.Dimension.Rows, worksheet.Dimension.Columns].Style.Fill.PatternType =
                ExcelFillStyle.Solid;
            worksheet.Cells[1, 1, worksheet.Dimension.Rows, worksheet.Dimension.Columns].Style.Fill.BackgroundColor
                .SetColor(Styles.BackgroundColor.General);
            worksheet.Cells[1, 1, worksheet.Dimension.Rows, worksheet.Dimension.Columns].Style.Font.Color
                .SetColor(Styles.FontColor.General);

            worksheet.Cells[1, 1, worksheet.Dimension.Rows, worksheet.Dimension.Columns].Style.Border.Top.Style =
                ExcelBorderStyle.Thin;
            worksheet.Cells[1, 1, worksheet.Dimension.Rows, worksheet.Dimension.Columns].Style.Border.Top.Color
                .SetColor(Styles.BorderColor.Cell);

            worksheet.Cells[1, 1, worksheet.Dimension.Rows, worksheet.Dimension.Columns].Style.Border.Bottom.Style =
                ExcelBorderStyle.Thin;
            worksheet.Cells[1, 1, worksheet.Dimension.Rows, worksheet.Dimension.Columns].Style.Border.Bottom.Color
                .SetColor(Styles.BorderColor.Cell);

            worksheet.Cells[1, 1, worksheet.Dimension.Rows, worksheet.Dimension.Columns].Style.Border.Left.Style =
                ExcelBorderStyle.Thin;
            worksheet.Cells[1, 1, worksheet.Dimension.Rows, worksheet.Dimension.Columns].Style.Border.Left.Color
                .SetColor(Styles.BorderColor.Cell);

            worksheet.Cells[1, 1, worksheet.Dimension.Rows, worksheet.Dimension.Columns].Style.Border.Right.Style =
                ExcelBorderStyle.Thin;
            worksheet.Cells[1, 1, worksheet.Dimension.Rows, worksheet.Dimension.Columns].Style.Border.Right.Color
                .SetColor(Styles.BorderColor.Cell);
        }

        /// <summary>
        ///     Automatic resize all columns
        /// </summary>
        /// <param name="worksheet"></param>
        public static void AutoFitAllCol(ExcelWorksheet worksheet)
        {
            for (var x = 1; x <= worksheet.Dimension.End.Column; x++) worksheet.Column(x).AutoFit();
        }

        /// <summary>
        ///     Method to lock a line
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="iRow"></param>
        public static void LockRow(ExcelWorksheet worksheet, int iRow)
        {
            worksheet.Row(iRow).Style.Locked = true;
            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Fill.BackgroundColor
                .SetColor(Styles.BackgroundColor.CellLocked);
            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Font.Color
                .SetColor(Styles.FontColor.CellLocked);
        }

        /// <summary>
        ///     Unlocks all cells in a row
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="iRow"></param>
        public static void Unlock(ExcelWorksheet worksheet, int iRow)
        {
            worksheet.Row(iRow).Style.Locked = false;
            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Fill.BackgroundColor
                .SetColor(Styles.BackgroundColor.CellUnlocked);
            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Font.Color
                .SetColor(Styles.FontColor.CellUnlocked);
        }

        /// <summary>
        ///     Unlocks a cell
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="iRow"></param>
        /// <param name="iCol"></param>
        public static void UnlockCell(ExcelWorksheet worksheet, int iRow, int iCol)
        {
            worksheet.Cells[iRow, iCol].Style.Locked = false;
            worksheet.Cells[iRow, iCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[iRow, iCol].Style.Fill.BackgroundColor.SetColor(Styles.BackgroundColor.CellUnlocked);
            worksheet.Cells[iRow, iCol].Style.Font.Color.SetColor(Styles.FontColor.CellUnlocked);
        }

        /// <summary>
        ///     Formats the line that contains the value types
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="iRow"></param>
        public static void FormattingTheHeaderTypeField(ExcelWorksheet worksheet, int iRow)
        {
            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Fill.BackgroundColor
                .SetColor(Styles.BackgroundColor.HeaderTypeField);
            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Font.Color
                .SetColor(Styles.FontColor.HeaderTypeField);
            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Font.Italic = true;

            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Border.Left.Color
                .SetColor(Styles.BorderColor.HeaderTypeFieldCell);

            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Border.Right.Style =
                ExcelBorderStyle.Thin;
            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Border.Right.Color
                .SetColor(Styles.BorderColor.HeaderTypeFieldCell);

            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Border.Bottom.Style =
                ExcelBorderStyle.Medium;
            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Border.Bottom.Color
                .SetColor(Styles.BorderColor.HeaderTypeField);

            worksheet.Cells[iRow, 1].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[iRow, 1].Style.Border.Left.Color.SetColor(Styles.BorderColor.HeaderTypeField);

            worksheet.Cells[iRow, worksheet.Dimension.Columns].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[iRow, worksheet.Dimension.Columns].Style.Border.Right.Color
                .SetColor(Styles.BorderColor.HeaderTypeField);
        }

        /// <summary>
        ///     Formats a line of subtotal level 1
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="iRow"></param>
        public static void FormatingLevel1(ExcelWorksheet worksheet, int iRow)
        {
            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Locked = true;
            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Font.Bold = true;
            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Font.Color
                .SetColor(Styles.FontColor.Level1);
            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Fill.BackgroundColor
                .SetColor(Styles.BackgroundColor.Level1);

            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Border.Top.Style =
                ExcelBorderStyle.Medium;
            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Border.Top.Color
                .SetColor(Styles.BorderColor.Level1);

            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Border.Bottom.Style =
                ExcelBorderStyle.Medium;
            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Border.Bottom.Color
                .SetColor(Styles.BorderColor.Level1);

            worksheet.Cells[iRow, 1].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[iRow, 1].Style.Border.Left.Color.SetColor(Styles.BorderColor.Level1);

            worksheet.Cells[iRow, worksheet.Dimension.Columns].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[iRow, worksheet.Dimension.Columns].Style.Border.Right.Color
                .SetColor(Styles.BorderColor.Level1);
        }

        /// <summary>
        ///     Formats a line of subtotal level 2
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="iRow"></param>
        public static void FormatingLevel2(ExcelWorksheet worksheet, int iRow)
        {
            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Locked = true;
            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Font.Bold = true;
            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Font.Color
                .SetColor(Styles.FontColor.Level2);
            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Fill.BackgroundColor
                .SetColor(Styles.BackgroundColor.Level2);

            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Border.Top.Style =
                ExcelBorderStyle.Medium;
            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Border.Top.Color
                .SetColor(Styles.BorderColor.Level2);

            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Border.Bottom.Style =
                ExcelBorderStyle.Medium;
            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Border.Bottom.Color
                .SetColor(Styles.BorderColor.Level2);

            worksheet.Cells[iRow, 1].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[iRow, 1].Style.Border.Left.Color.SetColor(Styles.BorderColor.Level2);

            worksheet.Cells[iRow, worksheet.Dimension.Columns].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[iRow, worksheet.Dimension.Columns].Style.Border.Right.Color
                .SetColor(Styles.BorderColor.Level2);
        }

        /// <summary>
        ///     Formats a line of total
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="iRow"></param>
        public static void FormatingTotalRow(ExcelWorksheet worksheet, int iRow)
        {
            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Locked = true;
            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Font.Bold =
                true; //Active le caractère gras
            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Font.Color
                .SetColor(Styles.FontColor.Total);
            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Fill.BackgroundColor
                .SetColor(Styles.BackgroundColor.Total);

            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Border.Top.Style = ExcelBorderStyle.Thick;
            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Border.Top.Color
                .SetColor(Styles.BorderColor.Total);

            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Border.Bottom.Style =
                ExcelBorderStyle.Thick;
            worksheet.Cells[iRow, 1, iRow, worksheet.Dimension.Columns].Style.Border.Bottom.Color
                .SetColor(Styles.BorderColor.Total);

            worksheet.Cells[iRow, 1].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[iRow, 1].Style.Border.Left.Color.SetColor(Styles.BorderColor.Total);

            worksheet.Cells[iRow, worksheet.Dimension.Columns].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[iRow, worksheet.Dimension.Columns].Style.Border.Right.Color
                .SetColor(Styles.BorderColor.Total);
        }

        /// <summary>
        ///     Freezes a row of the Excel table
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="iRow"></param>
        public static void FreezeRow(ExcelWorksheet worksheet, int iRow)
        {
            worksheet.View.FreezePanes(iRow, 1);
        }


        /// <summary>
        ///     Gets the operator in string format according to its type
        /// </summary>
        /// <param name="sFilterType"></param>
        /// <returns>Returns the operator in string format or an empty string</returns>
        private static string GetOperator(ScheduleFilterType sFilterType)
        {
            var sResult = string.Empty;
            switch (sFilterType)
            {
                case ScheduleFilterType.Equal:
                    sResult = "=";
                    break;
                case ScheduleFilterType.LessThan:
                    sResult = "<";
                    break;
                case ScheduleFilterType.LessThanOrEqual:
                    sResult = "<=";
                    break;
                case ScheduleFilterType.GreaterThan:
                    sResult = ">";
                    break;
                case ScheduleFilterType.GreaterThanOrEqual:
                    sResult = ">=";
                    break;
                case ScheduleFilterType.NotEqual:
                    sResult = "<>";
                    break;
            }

            return sResult;
        }

        /// <summary>
        ///     Gets the string for sorting
        /// </summary>
        /// <param name="vSchedule">ViewSchedule</param>
        /// <param name="doc">Document</param>
        /// <param name="FieldsList">List<ScheduleField></param>
        /// <returns>Returns the string for sorting or an empty string</returns>
        public static string GetStringSort(ViewSchedule vSchedule, Document doc, List<ScheduleField> FieldsList)
        {
            var sStringSort = string.Empty;
            foreach (var scheduleSortGroupField in vSchedule.Definition.GetSortGroupFields())
            {
                var scheduleField = FieldsList.FirstOrDefault(f => f.FieldId == scheduleSortGroupField.FieldId);
                if (scheduleField != null)
                {
                    if (sStringSort != string.Empty) sStringSort += ", ";
                    sStringSort += string.Format("[{0}] {1}", scheduleField.GetSchedulableField().GetName(doc),
                        scheduleSortGroupField.SortOrder == ScheduleSortOrder.Ascending ? "asc" : "desc");
                }
            }

            return sStringSort;
        }

        /// <summary>
        ///     Gets the string to filter the table
        /// </summary>
        /// <param name="vSchedule">ViewSchedule</param>
        /// <returns>Returns the string to filter or an empty string</returns>
        public static string GetStringFilter(ViewSchedule vSchedule, Document doc)
        {
            var sStringFilter = string.Empty;
            foreach (var sFilter in vSchedule.Definition.GetFilters())
            {
                var sOperator = GetOperator(sFilter.FilterType);
                var sField = vSchedule.Definition.GetField(sFilter.FieldId);
                if (sField != null && !string.IsNullOrEmpty(sOperator))
                {
                    var bValueFind = false;
                    var sFilterItem = "[" + sField.GetName() + "]" + sOperator;
                    if (sFilter.IsDoubleValue)
                    {
#if REVIT2021
                            var foItem = doc.GetUnits().GetFormatOptions(sField.GetSpecTypeId());
                            var dut = foItem.GetUnitTypeId();
                            double dValue = UnitUtils.ConvertFromInternalUnits(sFilter.GetDoubleValue(), dut);
                            sFilterItem += dValue.ToString();
#else
                        var foItem = doc.GetUnits().GetFormatOptions(sField.UnitType);
                        var dut = foItem.DisplayUnits;
                        var dValue = UnitUtils.ConvertFromInternalUnits(sFilter.GetDoubleValue(), dut);
                        sFilterItem += dValue.ToString();
#endif
                        bValueFind = true;
                    }
                    else if (sFilter.IsIntegerValue)
                    {
                        sFilterItem += "'" + sFilter.GetIntegerValue() + "'";
                        bValueFind = true;
                    }
                    else if (sFilter.IsStringValue)
                    {
                        sFilterItem += "'" + sFilter.GetStringValue().AddSlashes() + "'";
                        bValueFind = true;
                    }
                    else if (sFilter.IsElementIdValue)
                    {
                        sFilterItem += sFilter.GetElementIdValue().ToString();
                    }
                    else if (sFilter.IsNullValue)
                    {
                        sFilterItem += "null";
                    }

                    if (bValueFind)
                    {
                        if (sStringFilter != string.Empty) sStringFilter += " and ";
                        sStringFilter += sFilterItem;
                    }
                }
            }

            return sStringFilter;
        }

        public static string GetElementFamilyName(Document doc, ElementType elementType)
        {
            return elementType.FamilyName;
        }

        public static Parameter GetElementParameter(Element assetElement, string parameterName)
        {
            if (assetElement == null) return null;

            return assetElement.LookupParameter(parameterName);
        }
    }
}