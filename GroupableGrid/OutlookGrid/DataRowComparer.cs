using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

/// <summary>
/// reusable custom DataRow comparer implementation, can be used to sort DataTables
/// </summary>
public class DataRowComparer : IComparer
{
	ListSortDirection direction;
	int columnIndex;

	public DataRowComparer(int columnIndex, ListSortDirection direction)
	{
		this.columnIndex = columnIndex;
		this.direction = direction;
	}

	#region IComparer Members

	public int Compare(object x, object y)
	{

		DataRow obj1 = (DataRow)x;
		DataRow obj2 = (DataRow)y;
		return string.Compare(obj1[columnIndex].ToString(), obj2[columnIndex].ToString()) * (direction == ListSortDirection.Ascending ? 1 : -1);
	}

	#endregion
}