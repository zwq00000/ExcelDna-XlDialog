using System;
using ExcelDna.Integration;

namespace ExcelDNA.XlDialogs
{

    internal static class Extensions
    {

        public static bool IsNull(this object instance)
        {
            return instance == null || instance is DBNull || instance is ExcelEmpty || instance is ExcelError ||
                        instance is ExcelMissing
                        || instance == System.Type.Missing;
        }

    }
}