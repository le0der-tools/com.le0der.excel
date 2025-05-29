using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using System;

namespace Le0der.Toolkits.Excel
{
	[AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = false)]
	public class ExcelAssetAttribute : Attribute
	{
		public string AssetPath { get; set; }
		public string ExcelName { get; set; }
		public bool LogOnImport { get; set; }
	}
}