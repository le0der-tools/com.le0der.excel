using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using System;

namespace Le0der.Toolkits.Excel
{
	[AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = false)]
	public class ExcelAssetAttribute : Attribute
	{
		public bool IsRelative { get; set; }        //  是否相对路径
		public string AssetPath { get; set; }       //  生成的资源路径
		public string ExcelName { get; set; }       //  Excel名称（用于处理Excel表格名称和脚本名称不匹配的问题）
		public bool LogOnImport { get; set; }       //  是否在导入Excel时打印日志
	}
}