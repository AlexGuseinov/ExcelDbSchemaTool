using ExcelDbSchemaTool.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDbSchemaTool;

public static class ListOfTableDBModelExtensions
{
    public static List<Table> ToTables(this List<TableDBModel> tableFromDBs)
    {
        List<Table> result = new();
        var groupByed = tableFromDBs.GroupBy(x => x.TableName);

        foreach (var columnsInGroup in groupByed)
        {
            Table table = new();
            table.Name = columnsInGroup.Key;

            foreach (var column in columnsInGroup)
            {
                Column columnDetail = new();
                columnDetail.Name = column.ColumnName;
                columnDetail.Type = column.ColumnType;
                columnDetail.Description = column.Comment;
                table.Columns.Add(columnDetail);

            }
            result.Add(table);
        }

        return result;
    }
}
