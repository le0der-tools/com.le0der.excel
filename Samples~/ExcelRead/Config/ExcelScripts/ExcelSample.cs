using System;
using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using Le0der.Toolkits.Excel;

namespace Le0der.Toolkits.Excel.Demo
{
	[ExcelAsset(IsRelative = true, AssetPath = "../ExcelDatas")]
	public class ExcelSample : ScriptableObject
	{
		public List<SheetEntitySample> Sample; // Replace 'SheetEntitySample' to an actual type that is serializable.
		public List<SheetEntitySample2> Sample2; // Replace 'SheetEntitySample2' to an actual type that is serializable.
	}
}