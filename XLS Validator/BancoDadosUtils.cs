using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;

namespace AtosCapital
{
    public static class BancoDadosUtils
    {
        public static IList<T> ConvertDataTableToDataType<T>(DataTable dt)
        {
            List<T> result = new List<T>();

            if (dt.Rows.Count > 0)
            {
                Type typeT = typeof(T);
                Type nullableType = Nullable.GetUnderlyingType(typeT);
                if (typeT.IsPrimitive || typeT == typeof(string) || typeT == typeof(DateTime) || typeT == typeof(decimal) || typeT == typeof(double) || nullableType != null)
                {
                    // Tipo primitivo
                    foreach (DataRow row in dt.Rows)
                    {
                        T item = default(T);

                        if (!row.IsNull(0))
                        {
                            object readData = row[0];
                            if (readData is T)
                                item = (T)readData;
                            else
                                item = (T)Convert.ChangeType(readData, typeof(T));
                        }
                        result.Add(item);
                    }
                }
                else
                {
                    PropertyInfo[] props = typeT.GetProperties();
                    string[] colunas = dt.Columns.Cast<DataColumn>().Select(x => x.ColumnName).ToArray();
                    PropertyInfo[] propsCommon = props.Where(x => colunas.Contains(x.Name)).ToArray();

                    foreach (DataRow row in dt.Rows)
                    {
                        T item = Activator.CreateInstance<T>();
                        if (propsCommon.Length == 0 && colunas.Length == 1)
                        {
                            object readData = row[0];
                            if (readData is T)
                                item = (T)readData;
                            else
                                item = (T)Convert.ChangeType(readData, typeof(T));
                        }
                        else
                        {
                            foreach (PropertyInfo pro in propsCommon)
                            {
                                if (!row.IsNull(pro.Name))
                                    pro.SetValue(item, row[pro.Name], null);
                            }
                        }
                        result.Add(item);
                    }
                }
            }

            return result;
        }
    }
}
