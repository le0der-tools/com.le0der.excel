using System;
using System.Collections;
using System.Collections.Generic;
using UnityEngine;

namespace Le0der.Toolkits.Excel.Demo
{
	[ExcelAsset]
	public class ExcelMstItems : ScriptableObject
	{
		public List<SheetEntityEntity> Entity; // Replace 'SheetEntityEntity' to an actual type that is serializable.
	}
}

