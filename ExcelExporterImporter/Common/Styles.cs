using System.Drawing;

namespace ExcelExporterImporter.Common
{
    public static class Styles
    {
        public static class BackgroundColor
        {
            public static Color CellLocked = Color.FromArgb(220, 220, 220); //Cellule verouillé #DCDCDC
            public static Color CellUnlocked = Color.FromArgb(255, 255, 255); //Cellule déverouiller #FFFFFF
            public static Color clColumnTitle = Color.FromArgb(61, 61, 61); //3D3D3D
            public static Color clDescriptionRow = Color.FromArgb(255, 255, 255); //#FFFFFF
            public static Color clLegendTitle = Color.FromArgb(0, 0, 0); //#000000
            public static Color ColElementType = Color.FromArgb(186, 237, 255); //#BAEDFF
            public static Color DefaultLevel = Color.FromArgb(173, 173, 173); //Couleur par defaut #ADADAD
            public static Color General = Color.FromArgb(255, 255, 255); //FFFFFF
            public static Color Header = Color.FromArgb(59, 59, 59); //Entête / Header

            public static Color
                HeaderTypeField = Color.FromArgb(194, 225, 211); //Entête type / Header type field #C2E1D3

            public static Color Level1 = Color.FromArgb(143, 143, 143); //#8F8F8F;
            public static Color Level2 = Color.FromArgb(199, 199, 199); //#C7C7C7
            public static Color MsgNotBeImported = Color.FromArgb(0, 0, 0); //#000000
            public static Color Total = Color.FromArgb(77, 75, 75); //Cellule Total
            public static Color TypeFormula = Color.FromArgb(225, 158, 32); //#E19E20
        }

        public static class FontColor
        {
            public static Color CellLocked = Color.FromArgb(0, 0, 0); //Cellule verouillé E0E0E0
            public static Color CellUnlocked = Color.FromArgb(0, 0, 0); //Cellule déverouiller #000000.
            public static Color clColumnTitle = Color.FromArgb(255, 255, 255); //#FFFFFF
            public static Color clDescriptionRow = Color.FromArgb(0, 0, 0); //#000000
            public static Color clLegendTitle = Color.FromArgb(255, 255, 255); //#FFFFFF
            public static Color ColElementType = Color.FromArgb(0, 0, 0); //#000000
            public static Color DefaultLevel = Color.FromArgb(0, 0, 0); //Couleur par defaut pour un niveau #000000
            public static Color General = Color.FromArgb(0, 0, 0); //#000000
            public static Color Header = Color.FromArgb(255, 255, 255); //Entête / Header
            public static Color HeaderTypeField = Color.FromArgb(0, 0, 0); //Entête type / Header type field
            public static Color Level1 = Color.FromArgb(255, 255, 255); //FFFFFF
            public static Color Level2 = Color.FromArgb(0, 0, 0); //00000
            public static Color MsgNotBeImported = Color.FromArgb(255, 255, 255); //#FFFFFF
            public static Color Total = Color.FromArgb(255, 255, 255); //Cellule Total / Total cellule  
            public static Color TypeFormula = Color.FromArgb(0, 0, 0); //00000
        }

        public static class BorderColor
        {
            public static Color Cell = Color.FromArgb(0, 0, 0); //#000000
            public static Color clLegendTitle = Color.FromArgb(0, 0, 0); //#FFFFFF
            public static Color clColumnTitle = Color.FromArgb(0, 0, 0); //#000000
            public static Color clDescriptionRow = Color.FromArgb(0, 0, 0); //#000000
            public static Color Header = Color.FromArgb(0, 0, 0); //#000000
            public static Color HeaderCell = Color.FromArgb(255, 255, 255); //#FFFFFF
            public static Color HeaderTypeField = Color.FromArgb(0, 0, 0); //#000000
            public static Color HeaderTypeFieldCell = Color.FromArgb(0, 0, 0); //#000000
            public static Color Level1 = Color.FromArgb(0, 0, 0); //#000000
            public static Color Level2 = Color.FromArgb(0, 0, 0); //#000000
            public static Color Total = Color.FromArgb(0, 0, 0); //#000000
            public static Color TypeFormula = Color.FromArgb(0, 0, 0); //#000000
        }
    }
}