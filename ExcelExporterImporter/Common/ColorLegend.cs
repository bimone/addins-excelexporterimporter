using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExcelExporterImporter.Common
{
    internal static class ColorLegend
    {
        /// <summary>
        ///     Addition of the color legend in the legend tab
        /// </summary>
        /// <param name="Worksheet">Excel table</param>
        public static void Add(ExcelWorksheet Worksheet)
        {
            var iRow = 2;
            var iCol = 2;
            //Table title
            Worksheet.Cells[iRow, iCol].Value = Resources.clLegendTitle;
            FormatTitleLegend(Worksheet, iRow, iCol, iRow, iCol + 1);
            //Column titles
            iRow += 1;
            Worksheet.Cells[iRow, iCol].Value = Resources.clColTitleColor;
            Worksheet.Cells[iRow, iCol + 1].Value = Resources.clColTitleDescription;
            FormatTitleColumn(Worksheet, iRow, iCol, iRow, iCol + 1);

            //Description: Column title
            AddDescriptionRow(Worksheet, ref iRow, iCol, Resources.clDescColumnTitle, Styles.BackgroundColor.Header);

            //Description : Value title
            AddDescriptionRow(Worksheet, ref iRow, iCol, Resources.clDescValueType,
                Styles.BackgroundColor.HeaderTypeField);

            //Description : Parameter value locked
            AddDescriptionRow(Worksheet, ref iRow, iCol, Resources.clDescParameterValueLocked,
                Styles.BackgroundColor.CellLocked);

            //Description : Value of a locked element type parameter
            AddDescriptionRow(Worksheet, ref iRow, iCol, Resources.clDescValueLockedElementType,
                Styles.BackgroundColor.ColElementType);

            //Description : Value of a parameter which cannot be exported
            AddDescriptionRow(Worksheet, ref iRow, iCol, Resources.clDescValueParameterNotExported,
                Styles.BackgroundColor.TypeFormula);

            //Description : Changing the value of the parameter is allowed
            AddDescriptionRow(Worksheet, ref iRow, iCol, Resources.clDescChangingValueParameterAllowed,
                Styles.BackgroundColor.CellUnlocked);

            //Description : Sub-total level 1
            AddDescriptionRow(Worksheet, ref iRow, iCol, Resources.clDescSubTotalLevel1, Styles.BackgroundColor.Level1);

            //Couleur Niveau 2 et plus
            AddDescriptionRow(Worksheet, ref iRow, iCol, Resources.clDescSubTotalLevel2, Styles.BackgroundColor.Level2);

            //Couleur Total
            AddDescriptionRow(Worksheet, ref iRow, iCol, Resources.clDescTotal, Styles.BackgroundColor.Total);

            //Format Last Line
            FormatLastDescriptionLine(Worksheet, iRow, iCol, iRow, iCol + 1);
            AjustColumnWidth(Worksheet);
            //Insert Legend Key
            Worksheet.Cells[1, 1].Value = Constants.LegendUniqueId;
            Worksheet.Row(1).Hidden = true;
        }

        /// <summary>
        ///     Format the title area
        /// </summary>
        /// <param name="Worksheet"></param>
        /// <param name="iRow"></param>
        /// <param name="iCol"></param>
        /// <param name="iRow2"></param>
        /// <param name="iCol2"></param>
        private static void FormatTitleLegend(ExcelWorksheet Worksheet, int iRow, int iCol, int iRow2, int iCol2)
        {
            Worksheet.Cells[iRow, iCol, iRow2, iCol2].Merge = true;
            Worksheet.Cells[iRow, iCol, iRow2, iCol2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            Worksheet.Cells[iRow, iCol, iRow2, iCol2].Style.Locked = true;
            Worksheet.Cells[iRow, iCol, iRow2, iCol2].Style.Font.Bold = true;
            Worksheet.Cells[iRow, iCol, iRow2, iCol2].Style.Font.Size = 16;
            Worksheet.Cells[iRow, iCol, iRow2, iCol2].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Worksheet.Cells[iRow, iCol, iRow2, iCol2].Style.Font.Color.SetColor(Styles.FontColor.clLegendTitle);
            Worksheet.Cells[iRow, iCol, iRow2, iCol2].Style.Fill.BackgroundColor
                .SetColor(Styles.BackgroundColor.clLegendTitle);

            Worksheet.Cells[iRow, iCol, iRow2, iCol2].Style.Border.Top.Style = ExcelBorderStyle.Thick;
            Worksheet.Cells[iRow, iCol, iRow2, iCol2].Style.Border.Top.Color.SetColor(Styles.BorderColor.clLegendTitle);

            Worksheet.Cells[iRow, iCol, iRow2, iCol2].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
            Worksheet.Cells[iRow, iCol, iRow2, iCol2].Style.Border.Bottom.Color
                .SetColor(Styles.BorderColor.clLegendTitle);

            Worksheet.Cells[iRow, iCol].Style.Border.Left.Style = ExcelBorderStyle.Thick;
            Worksheet.Cells[iRow, iCol].Style.Border.Left.Color.SetColor(Styles.BorderColor.clLegendTitle);

            Worksheet.Cells[iRow, iCol2].Style.Border.Right.Style = ExcelBorderStyle.Thick;
            Worksheet.Cells[iRow, iCol2].Style.Border.Right.Color.SetColor(Styles.BorderColor.clLegendTitle);
        }

        /// <summary>
        ///     Format column headings
        /// </summary>
        /// <param name="Worksheet"></param>
        /// <param name="iRow"></param>
        /// <param name="iCol"></param>
        /// <param name="iRow2"></param>
        /// <param name="iCol2"></param>
        private static void FormatTitleColumn(ExcelWorksheet Worksheet, int iRow, int iCol, int iRow2, int iCol2)
        {
            Worksheet.Cells[iRow, iCol, iRow2, iCol2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            Worksheet.Cells[iRow, iCol, iRow2, iCol2].Style.Locked = true;
            Worksheet.Cells[iRow, iCol, iRow2, iCol2].Style.Font.Bold = true;
            Worksheet.Cells[iRow, iCol, iRow2, iCol2].Style.Font.Size = 14;
            Worksheet.Cells[iRow, iCol, iRow2, iCol2].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Worksheet.Cells[iRow, iCol, iRow2, iCol2].Style.Font.Color.SetColor(Styles.FontColor.clColumnTitle);
            Worksheet.Cells[iRow, iCol, iRow2, iCol2].Style.Fill.BackgroundColor
                .SetColor(Styles.BackgroundColor.clColumnTitle);

            Worksheet.Cells[iRow, iCol].Style.Border.Left.Style = ExcelBorderStyle.Thick;
            Worksheet.Cells[iRow, iCol].Style.Border.Left.Color.SetColor(Styles.BorderColor.clColumnTitle);

            Worksheet.Cells[iRow, iCol].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            Worksheet.Cells[iRow, iCol].Style.Border.Right.Color.SetColor(Styles.BorderColor.clColumnTitle);

            Worksheet.Cells[iRow, iCol2].Style.Border.Right.Style = ExcelBorderStyle.Thick;
            Worksheet.Cells[iRow, iCol2].Style.Border.Right.Color.SetColor(Styles.BorderColor.clColumnTitle);

            Worksheet.Cells[iRow, iCol, iRow2, iCol2].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
            Worksheet.Cells[iRow, iCol, iRow2, iCol2].Style.Border.Bottom.Color
                .SetColor(Styles.BorderColor.clColumnTitle);
        }

        private static void FormatDescriptionLine(ExcelWorksheet Worksheet, int iRow, int iCol, int iRow2, int iCol2)
        {
            Worksheet.Cells[iRow, iCol, iRow2, iCol2].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Worksheet.Cells[iRow2, iCol2].Style.Font.Color.SetColor(Styles.FontColor.clDescriptionRow);
            Worksheet.Cells[iRow2, iCol2].Style.Fill.BackgroundColor.SetColor(Styles.BackgroundColor.clDescriptionRow);

            Worksheet.Cells[iRow, iCol].Style.Border.Left.Style = ExcelBorderStyle.Thick;
            Worksheet.Cells[iRow, iCol].Style.Border.Left.Color.SetColor(Styles.BorderColor.clDescriptionRow);

            Worksheet.Cells[iRow, iCol].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            Worksheet.Cells[iRow, iCol].Style.Border.Right.Color.SetColor(Styles.BorderColor.clDescriptionRow);

            Worksheet.Cells[iRow, iCol2].Style.Border.Right.Style = ExcelBorderStyle.Thick;
            Worksheet.Cells[iRow, iCol2].Style.Border.Right.Color.SetColor(Styles.BorderColor.clDescriptionRow);

            Worksheet.Cells[iRow, iCol, iRow2, iCol2].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            Worksheet.Cells[iRow, iCol, iRow2, iCol2].Style.Border.Bottom.Color
                .SetColor(Styles.BorderColor.clDescriptionRow);

            Worksheet.Cells[iRow, iCol, iRow2, iCol2].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            Worksheet.Cells[iRow, iCol, iRow2, iCol2].Style.Border.Top.Color
                .SetColor(Styles.BorderColor.clDescriptionRow);
        }

        /// <summary>
        ///     Replaces bottom cell line with thicker line
        /// </summary>
        /// <param name="Worksheet"></param>
        /// <param name="iRow"></param>
        /// <param name="iCol"></param>
        /// <param name="iRow2"></param>
        /// <param name="iCol2"></param>
        private static void FormatLastDescriptionLine(ExcelWorksheet Worksheet, int iRow, int iCol, int iRow2,
            int iCol2)
        {
            Worksheet.Cells[iRow, iCol, iRow2, iCol2].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
            Worksheet.Cells[iRow, iCol, iRow2, iCol2].Style.Border.Bottom.Color
                .SetColor(Styles.BorderColor.clDescriptionRow);
        }

        /// <summary>
        ///     Automatically adjust column width
        /// </summary>
        /// <param name="Worksheet"></param>
        private static void AjustColumnWidth(ExcelWorksheet Worksheet)
        {
            for (var x = 1; x <= Worksheet.Dimension.Columns; x++) Worksheet.Column(x).AutoFit();
        }

        private static void AddDescriptionRow(ExcelWorksheet Worksheet, ref int iRow, int iCol, string sDescription,
            Color cColorCell)
        {
            iRow += 1;
            Worksheet.Cells[iRow, iCol].Value = "";
            Worksheet.Cells[iRow, iCol + 1].Value = sDescription;
            Worksheet.Cells[iRow, iCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Worksheet.Cells[iRow, iCol].Style.Fill.BackgroundColor.SetColor(cColorCell);
            FormatDescriptionLine(Worksheet, iRow, iCol, iRow, iCol + 1);
        }
    }
}