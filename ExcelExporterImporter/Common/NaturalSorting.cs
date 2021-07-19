using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace ExcelExporterImporter.Common
{
    internal class NaturalSorting
    {
        /// <summary>
        ///     Sorting a datatable using natural sorting
        /// </summary>
        /// <param name="dt">The sorting datatable</param>
        /// <param name="sSort">Sorting string</param>
        /// <param name="sMsgError">Error mesage</param>
        /// <returns>Returns the sort datatable or an empty datatable</returns>
        public static DataTable DataTableSort(DataTable dt, string sSort, out string sMsgError)
        {
            sMsgError = string.Empty;
            var dtNew = new DataTable();
            try
            {
                dtNew = dt.Clone();
                var ListSort = CreateSortClassList(dtNew, sSort, out sMsgError);
                if (string.IsNullOrEmpty(sMsgError))
                {
                    var iNumberRow = dt.Rows.Count;
                    //=====================================Test Tri par insertion==============================
                    //DataRow[] TabDataRow = new DataRow[iNumberRow];
                    for (var iIndex1 = 0; iIndex1 < iNumberRow; iIndex1++)
                    {
                        var InsertRow = dt.Rows[iIndex1];
                        if (iIndex1 == 0)
                        {
                            dtNew.Rows.Add(InsertRow.ItemArray);
                        }
                        else
                        {
                            var bFindPosition = false;
                            for (var iIndex2 = iIndex1 - 1; iIndex2 >= 0; iIndex2--)
                            {
                                var SecondRow = dtNew.Rows[iIndex2];
                                var ResultCompare = "wait";
                                foreach (var ItemSort in ListSort)
                                {
                                    ResultCompare = CompareTwoString(InsertRow[ItemSort.ColumnIndex].ToString(),
                                        SecondRow[ItemSort.ColumnIndex].ToString(), out sMsgError);
                                    if (string.IsNullOrEmpty(sMsgError))
                                    {
                                        if (ResultCompare == "less" && ItemSort.OrderBy == "desc" ||
                                            ResultCompare == "greater" && ItemSort.OrderBy == "asc")
                                        {
                                            bFindPosition = true;
                                            break;
                                        }

                                        if (ResultCompare != "egal") break;
                                    }
                                    else
                                    {
                                        throw new AddNewExeption(sMsgError);
                                    }
                                }

                                if (ResultCompare == "egal") bFindPosition = true;
                                if (bFindPosition)
                                {
                                    var iIndexPositionInsert = iIndex2 + 1;
                                    if (iIndexPositionInsert < iIndex1)
                                    {
                                        var NewDataRow = dtNew.NewRow();
                                        NewDataRow.ItemArray = InsertRow.ItemArray;
                                        dtNew.Rows.InsertAt(NewDataRow, iIndexPositionInsert);
                                    }
                                    else
                                    {
                                        dtNew.Rows.Add(InsertRow.ItemArray);
                                    }

                                    break;
                                }
                            }

                            if (!bFindPosition)
                            {
                                var NewDataRow = dtNew.NewRow();
                                NewDataRow.ItemArray = InsertRow.ItemArray;
                                dtNew.Rows.InsertAt(NewDataRow, 0);
                            }
                        }
                    }

                    //====================================Tri par sélection Actuel===============================
                    /*for (int iRow1 = 0;iRow1 < iNumberRow; iRow1++)
                    {
                        DataRow InsertRow = dt.Rows[0];
                        for(int iRow2 = 1;iRow2 < dt.Rows.Count;iRow2++)
                        {
                            DataRow SecondRow = dt.Rows[iRow2];
                            foreach(SortClass ItemSort in ListSort)
                            {
                                string ResultCompare = CompareTwoString(InsertRow[ItemSort.ColumnIndex].ToString(), SecondRow[ItemSort.ColumnIndex].ToString(),out sMsgError);
                                if(string.IsNullOrEmpty(sMsgError))
                                {
                                    if((ResultCompare == "less" && ItemSort.OrderBy == "desc") || (ResultCompare == "greater" && ItemSort.OrderBy == "asc"))
                                    {
                                        InsertRow = SecondRow;
                                        break;
                                    }
                                    else if(ResultCompare != "egal")
                                    {
                                        break;
                                    }
                                }
                                else
                                {
                                    throw new AddNewExeption(sMsgError);
                                }
                            }
                        }
                        dtNew.Rows.Add(InsertRow.ItemArray); //Add the row to the new table
                        dt.Rows.Remove(InsertRow); //Delete row from initial table
                    }*/
                }
                else
                {
                    throw new AddNewExeption(sMsgError);
                }
            }
            catch (Exception ex)
            {
                sMsgError = ex.Message;
            }

            return dtNew;
        }

        /// <summary>
        ///     Create a list from the sort string
        /// </summary>
        /// <param name="dt">Datatable</param>
        /// <param name="sSort">Sorting string</param>
        /// <returns>Return the list or an empty list</returns>
        private static List<SortClass> CreateSortClassList(DataTable dt, string sSort, out string sMsgError)
        {
            var ListClass = new List<SortClass>();
            sMsgError = string.Empty;
            try
            {
                char[] Separator = {','};
                char[] Separator2 = {' '};
                var TableOrder = sSort.Split(Separator);
                foreach (var sElementOrder in TableOrder)
                {
                    var ItemClass = new SortClass();
                    var iStartIndex = sElementOrder.IndexOf('[');
                    var iEndIndex = sElementOrder.IndexOf(']');
                    ItemClass.ColunmName = sElementOrder.Substring(iStartIndex + 1, iEndIndex - iStartIndex - 1);
                    if (sElementOrder.ToLower().Contains("desc"))
                        ItemClass.OrderBy = "desc";
                    else
                        ItemClass.OrderBy = "asc";
                    //We are looking for the index linked to the column
                    ItemClass.ColumnIndex = 0;
                    var bColumnFound = false;
                    for (var iCol = 0; iCol < dt.Columns.Count; iCol++)
                        if (dt.Columns[iCol].ColumnName.Trim().ToLower() == ItemClass.ColunmName.ToLower())
                        {
                            ItemClass.ColumnIndex = iCol;
                            bColumnFound = true;
                            break;
                        }

                    if (bColumnFound)
                        ListClass.Add(ItemClass);
                    else
                        throw new AddNewExeption(
                            string.Format("Cannot find index for column {0}", ItemClass.ColunmName));
                }
            }
            catch (Exception ex)
            {
                sMsgError = ex.Message;
            }

            return ListClass;
        }

        /// <summary>
        ///     Indicates if the string1 is smaller or equal or larger than the string2 by considering the numbers as numbers
        /// </summary>
        /// <param name="sString1">First string</param>
        /// <param name="sString2">Second string</param>
        /// <returns>Return less or egal or greater or error</returns>
        private static string CompareTwoString(string sString1, string sString2, out string sMsgError)
        {
            sMsgError = string.Empty;
            var sResult = "egal";
            try
            {
                if (!string.IsNullOrEmpty(sString1) && !string.IsNullOrEmpty(sString2))
                {
                    var ListString1 = GetSeparateLettersAndNumbersList(sString1);
                    var ListString2 = GetSeparateLettersAndNumbersList(sString2);
                    var iMaxChar = ListString1.Count() <= ListString2.Count()
                        ? ListString1.Count()
                        : ListString2.Count();
                    var x = 0;
                    while (x < iMaxChar && sResult == "egal")
                    {
                        var sChar1 = ListString1[x];
                        var sChar2 = ListString2[x];
                        if (IsNumber(sChar1) && IsNumber(sChar2))
                        {
                            if (Convert.ToDecimal(sChar1) < Convert.ToDecimal(sChar2))
                                sResult = "less";
                            else if (Convert.ToDecimal(sChar1) > Convert.ToDecimal(sChar2)) sResult = "greater";
                        }
                        else if (IsNumber(sChar1))
                        {
                            sResult = "less";
                        }
                        else if (IsNumber(sChar2))
                        {
                            sResult = "greater";
                        }
                        else
                        {
                            var bTabChar1 = Encoding.ASCII.GetBytes(sChar1.ToLower());
                            var bTabChar2 = Encoding.ASCII.GetBytes(sChar2.ToLower());
                            if (bTabChar1[0] < bTabChar2[0])
                                sResult = "less";
                            else if (bTabChar1[0] > bTabChar2[0]) sResult = "greater";
                        }

                        x += 1;
                    }
                }
                else if (!string.IsNullOrEmpty(sString1))
                {
                    sResult = "greater";
                }
                else if (!string.IsNullOrEmpty(sString2))
                {
                    sResult = "less";
                }
            }
            catch (Exception ex)
            {
                sMsgError = ex.Message;
            }

            return sResult;
        }

        /// <summary>
        ///     Separates numbers and letters from the string in a list
        /// </summary>
        /// <param name="sString">String to be separated</param>
        /// <returns>Returns the list with the result</returns>
        private static List<string> GetSeparateLettersAndNumbersList(string sString)
        {
            var ListItem = new List<string>();
            var sTemp = string.Empty;
            for (var x = 0; x < sString.Length; x++)
            {
                var sChr = sString.Substring(x, 1);
                if (IsNumber(sChr))
                {
                    sTemp += sChr;
                }
                else
                {
                    if (!string.IsNullOrEmpty(sTemp))
                    {
                        if (sChr == ".") //Possibility of a decimal
                        {
                            sTemp += sChr;
                        }
                        else
                        {
                            //We validate if the last character is a point
                            if (ValidateLastCharacterPoint(ref sTemp))
                            {
                                ListItem.Add(sTemp);
                                ListItem.Add(".");
                            }
                            else
                            {
                                ListItem.Add(sTemp);
                            }

                            sTemp = string.Empty;
                        }
                    }
                    else
                    {
                        ListItem.Add(sChr);
                    }
                }
            }

            if (!string.IsNullOrEmpty(sTemp))
            {
                //Valid if the last character is a point
                if (ValidateLastCharacterPoint(ref sTemp))
                {
                    ListItem.Add(sTemp);
                    ListItem.Add(".");
                }
                else
                {
                    ListItem.Add(sTemp);
                }
            }

            return ListItem;
        }

        /// <summary>
        ///     Valid if the last character is a point
        /// </summary>
        /// <param name="sString">String to validate</param>
        /// <returns>True or False and the string without the point in reference</returns>
        private static bool ValidateLastCharacterPoint(ref string sString)
        {
            var bResult = false;
            if (sString.Substring(sString.Length - 1, 1) == ".")
            {
                bResult = true;
                sString = sString.Substring(0, sString.Length - 1);
            }

            return bResult;
        }

        /// <summary>
        ///     Indicates whether the character sent is a number
        /// </summary>
        /// <param name="sChr"></param>
        /// <returns>True or False</returns>
        private static bool IsNumber(string sChr)
        {
            var bResult = false;
            int iNumber;
            decimal decNumber;
            double dNumber;
            if (int.TryParse(sChr, out iNumber) || decimal.TryParse(sChr, out decNumber) ||
                double.TryParse(sChr, out dNumber)) bResult = true;
            return bResult;
        }

        private class SortClass
        {
            public string ColunmName { get; set; }
            public string OrderBy { get; set; }
            public int ColumnIndex { get; set; }
        }

        private class AddNewExeption : Exception
        {
            public AddNewExeption(string message) : base(message)
            {
            }
        }
    }
}