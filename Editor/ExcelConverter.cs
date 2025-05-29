using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityEditor;
using System;
using System.IO;
using System.Reflection;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

namespace Le0der.Toolkits.Excel
{
	public class ExcelConverter : AssetPostprocessor
	{
		class ExcelAssetInfo
		{
			public Type AssetType { get; set; }
			public ExcelAssetAttribute Attribute { get; set; }
			public string ExcelName
			{
				get
				{
					return string.IsNullOrEmpty(Attribute.ExcelName) ? AssetType.Name : Attribute.ExcelName;
				}
			}
		}

		static List<ExcelAssetInfo> cachedInfos = null; // Clear on compile.

		static void OnPostprocessAllAssets(string[] importedAssets, string[] deletedAssets, string[] movedAssets, string[] movedFromAssetPaths)
		{
			bool imported = false;
			foreach (string path in importedAssets)
			{
				if (Path.GetExtension(path) == ".xls" || Path.GetExtension(path) == ".xlsx")
				{
					if (cachedInfos == null) cachedInfos = FindExcelAssetInfos();

					var excelName = Path.GetFileNameWithoutExtension(path);
					if (excelName.StartsWith("~$")) continue;

					string className = GetScriptClassString(excelName);
					ExcelAssetInfo info = cachedInfos.Find(i => i.ExcelName == className);

					if (info == null) continue;

					ImportExcel(path, info);
					imported = true;
				}
			}

			if (imported)
			{
				AssetDatabase.SaveAssets();
				AssetDatabase.Refresh();
			}
		}

		static List<ExcelAssetInfo> FindExcelAssetInfos()
		{
			var list = new List<ExcelAssetInfo>();
			foreach (var assembly in AppDomain.CurrentDomain.GetAssemblies())
			{
				foreach (var type in assembly.GetTypes())
				{
					var attributes = type.GetCustomAttributes(typeof(ExcelAssetAttribute), false);
					if (attributes.Length == 0) continue;
					var attribute = (ExcelAssetAttribute)attributes[0];
					var info = new ExcelAssetInfo()
					{
						AssetType = type,
						Attribute = attribute
					};
					list.Add(info);
				}
			}
			return list;
		}

		static UnityEngine.Object LoadOrCreateAsset(string assetPath, Type assetType)
		{
			Directory.CreateDirectory(Path.GetDirectoryName(assetPath));

			var asset = AssetDatabase.LoadAssetAtPath(assetPath, assetType);

			if (asset == null)
			{
				asset = ScriptableObject.CreateInstance(assetType.Name);
				AssetDatabase.CreateAsset((ScriptableObject)asset, assetPath);
				asset.hideFlags = HideFlags.NotEditable;
			}

			return asset;
		}

		static IWorkbook LoadBook(string excelPath)
		{
			using (FileStream stream = File.Open(excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				if (Path.GetExtension(excelPath) == ".xls") return new HSSFWorkbook(stream);
				else return new XSSFWorkbook(stream);
			}
		}

		static List<string> GetFieldNamesFromSheetHeader(ISheet sheet)
		{
			IRow headerRow = sheet.GetRow(0);

			var fieldNames = new List<string>();
			for (int i = 0; i < headerRow.LastCellNum; i++)
			{
				var cell = headerRow.GetCell(i);
				if (cell == null || cell.CellType == CellType.Blank) break;
				fieldNames.Add(cell.StringCellValue);
			}
			return fieldNames;
		}

		private static object CellToFieldObject(ICell cell, FieldInfo fieldInfo)
		{
			Type targetType = fieldInfo.FieldType;
			string cellValue = cell?.ToString() ?? string.Empty; // 获取单元格的字符串值

			try
			{
				if (targetType == typeof(string))
				{
					return cellValue; // 直接返回字符串
				}
				else if (targetType.IsEnum)
				{
					return Enum.Parse(targetType, cellValue); // 解析为枚举
				}
				else if (targetType == typeof(int))
				{
					return int.Parse(cellValue);
				}
				else if (targetType == typeof(float))
				{
					return float.Parse(cellValue);
				}
				else if (targetType == typeof(double))
				{
					return double.Parse(cellValue);
				}
				else if (targetType == typeof(bool))
				{
					// 支持多种布尔值表示方式
					if (cellValue.Equals("true", StringComparison.OrdinalIgnoreCase) || cellValue == "1")
						return true;
					else if (cellValue.Equals("false", StringComparison.OrdinalIgnoreCase) || cellValue == "0")
						return false;
					else
						throw new FormatException($"Invalid boolean value: {cellValue}");
				}
				else if (targetType == typeof(DateTime))
				{
					return DateTime.Parse(cellValue);
				}
				else if (targetType == typeof(TimeSpan))
				{
					return TimeSpan.Parse(cellValue);
				}
				else if (targetType == typeof(Guid))
				{
					return Guid.Parse(cellValue);
				}
				else if (targetType == typeof(decimal))
				{
					return decimal.Parse(cellValue);
				}
				else if (targetType == typeof(char))
				{
					return cellValue[0]; // 取第一个字符
				}
				else if (targetType == typeof(byte))
				{
					return byte.Parse(cellValue);
				}
				else if (targetType == typeof(short))
				{
					return short.Parse(cellValue);
				}
				else if (targetType == typeof(long))
				{
					return long.Parse(cellValue);
				}
				else if (targetType == typeof(Vector2))
				{
					// 假设单元格数据格式为 "x,y"
					string[] parts = cellValue.Split(',');
					if (parts.Length == 2)
						return new Vector2(float.Parse(parts[0]), float.Parse(parts[1]));
					else
						throw new FormatException($"Invalid Vector2 format: {cellValue}");
				}
				else if (targetType == typeof(Vector3))
				{
					// 假设单元格数据格式为 "x,y,z"
					string[] parts = cellValue.Split(',');
					if (parts.Length == 3)
						return new Vector3(float.Parse(parts[0]), float.Parse(parts[1]), float.Parse(parts[2]));
					else
						throw new FormatException($"Invalid Vector3 format: {cellValue}");
				}
				else if (targetType == typeof(Vector4))
				{
					// 假设单元格数据格式为 "x,y,z,w"
					string[] parts = cellValue.Split(',');
					if (parts.Length == 4)
						return new Vector4(float.Parse(parts[0]), float.Parse(parts[1]), float.Parse(parts[2]), float.Parse(parts[3]));
					else
						throw new FormatException($"Invalid Vector4 format: {cellValue}");
				}
				else
				{
					// 如果是值类型，返回默认值
					if (targetType.IsValueType)
						return Activator.CreateInstance(targetType);
					else
						return null; // 引用类型返回 null
				}
			}
			catch (Exception ex)
			{
				throw new Exception($"Failed to convert cell value to {targetType.Name}. Cell value: {cellValue}", ex);
			}
		}

		// 根据excel单元格类型，将单元格值转换为对应字段类型
		private static object CellToFieldObject(ICell cell, FieldInfo fieldInfo, bool isFormulaEvalute = false)
		{
			var type = isFormulaEvalute ? cell.CachedFormulaResultType : cell.CellType;

			switch (type)
			{
				case CellType.String:
					if (fieldInfo.FieldType.IsEnum)
						return Enum.Parse(fieldInfo.FieldType, cell.StringCellValue);
					else if (fieldInfo.FieldType == typeof(DateTime))
						return DateTime.Parse(cell.StringCellValue);
					else if (fieldInfo.FieldType == typeof(TimeSpan))
						return TimeSpan.Parse(cell.StringCellValue);
					else if (fieldInfo.FieldType == typeof(Guid))
						return Guid.Parse(cell.StringCellValue);
					else
						return cell.StringCellValue;
				case CellType.Boolean:
					return cell.BooleanCellValue;
				case CellType.Numeric:
					return Convert.ChangeType(cell.NumericCellValue, fieldInfo.FieldType);
				case CellType.Formula:
					if (isFormulaEvalute) return null;
					return CellToFieldObject(cell, fieldInfo, true);
				default:
					if (fieldInfo.FieldType.IsValueType)
					{
						return Activator.CreateInstance(fieldInfo.FieldType);
					}
					return null;
			}
		}

		static object CreateEntityFromRow(IRow row, List<string> columnNames, Type entityType, string sheetName)
		{
			var entity = Activator.CreateInstance(entityType);

			for (int i = 0; i < columnNames.Count; i++)
			{
				FieldInfo entityField = entityType.GetField(
					columnNames[i],
					BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic
				);
				if (entityField == null) continue;
				if (!entityField.IsPublic && entityField.GetCustomAttributes(typeof(SerializeField), false).Length == 0) continue;

				ICell cell = row.GetCell(i);
				if (cell == null) continue;

				try
				{
					object fieldValue = CellToFieldObject(cell, entityField);
					entityField.SetValue(entity, fieldValue);
				}
				catch
				{
					throw new Exception(string.Format("Invalid excel cell type at row {0}, column {1}, {2} sheet.", row.RowNum, cell.ColumnIndex, sheetName));
				}
			}
			return entity;
		}

		static object GetEntityListFromSheet(ISheet sheet, Type entityType)
		{
			List<string> excelColumnNames = GetFieldNamesFromSheetHeader(sheet);

			Type listType = typeof(List<>).MakeGenericType(entityType);
			MethodInfo listAddMethod = listType.GetMethod("Add", new Type[] { entityType });
			object list = Activator.CreateInstance(listType);

			// row of index 0 is header，row of index 1 is data type
			for (int i = 2; i <= sheet.LastRowNum; i++)
			{
				IRow row = sheet.GetRow(i);
				if (row == null) break;

				ICell entryCell = row.GetCell(0);
				if (entryCell == null || entryCell.CellType == CellType.Blank) break;

				// skip comment row
				if (entryCell.CellType == CellType.String && entryCell.StringCellValue.StartsWith("#")) continue;

				var entity = CreateEntityFromRow(row, excelColumnNames, entityType, sheet.SheetName);
				listAddMethod.Invoke(list, new object[] { entity });
			}
			return list;
		}

		static void ImportExcel(string excelPath, ExcelAssetInfo info)
		{
			string assetPath = "";
			string assetName = info.AssetType.Name + ".asset";

			if (string.IsNullOrEmpty(info.Attribute.AssetPath))
			{
				string basePath = Path.GetDirectoryName(excelPath);
				assetPath = Path.Combine(basePath, assetName);
			}
			else
			{
				var path = Path.Combine("Assets", info.Attribute.AssetPath);
				assetPath = Path.Combine(path, assetName);
			}
			UnityEngine.Object asset = LoadOrCreateAsset(assetPath, info.AssetType);

			IWorkbook book = LoadBook(excelPath);

			var assetFields = info.AssetType.GetFields();
			int sheetCount = 0;

			foreach (var assetField in assetFields)
			{
				ISheet sheet = book.GetSheet(assetField.Name);
				if (sheet == null) continue;

				Type fieldType = assetField.FieldType;
				if (!fieldType.IsGenericType || (fieldType.GetGenericTypeDefinition() != typeof(List<>))) continue;

				Type[] types = fieldType.GetGenericArguments();
				Type entityType = types[0];

				object entities = GetEntityListFromSheet(sheet, entityType);
				assetField.SetValue(asset, entities);
				sheetCount++;
			}

			if (info.Attribute.LogOnImport)
			{
				Debug.Log(string.Format("Imported {0} sheets form {1}.", sheetCount, excelPath));
			}

			EditorUtility.SetDirty(asset);
		}


		// 表格代码类名规范
		const string ScriptNameFormat = "Excel{0}";
		/// <summary>
		/// 获取Excel表格对应类名
		/// </summary>
		/// <param name="excelName">表格名称</param>
		/// <returns>类名</returns>
		public static string GetScriptClassString(string excelName)
		{
			return string.Format(ScriptNameFormat, excelName);
		}
	}
}