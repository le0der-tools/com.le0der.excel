using System.Collections.Generic;
using UnityEngine;
using UnityEditor;
using System.IO;
using System;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.Text;

namespace Le0der.Toolkits.Excel
{
	public class ExcelAssetScriptMenu
	{
		// 表格代码模板文件名称
		const string ScriptTemplateName = "ExcelAssetScriptTemplete.cs.txt";
		// 表格Sheet代码模板文件名称
		const string SheetScriptTemplateName = "ExcelAssetScriptEntityTemplete.cs.txt";
		// 表格Sheet代码类名规范
		const string SheetScriptNameFormat = "SheetEntity{0}";

		// 读取表格sheet的字段模板
		const string FieldTemplete = "\tpublic List<#EntityType#> #FIELDNAME#; // Replace '#EntityType#' to an actual type that is serializable.";

		/// <summary>
		/// 根据excel表格创建对应的数据代码
		/// </summary>
		public static void CreateScript()
		{
			// 保存路径
			string savePath = EditorUtility.SaveFolderPanel("Save ExcelAssetScript", Application.dataPath, "");
			if (savePath == "") return;

			// 选中的Excel文件
			var selectedAssets = Selection.GetFiltered(typeof(UnityEngine.Object), SelectionMode.Assets);

			// 获取Excel表格名称和sheet名称
			string excelPath = AssetDatabase.GetAssetPath(selectedAssets[0]);
			string excelName = Path.GetFileNameWithoutExtension(excelPath);
			List<ISheet> sheets = GetSheets(excelPath);

			// 获取对象代码的字符串
			List<string> scriptStrings = BuildScriptStrings(excelName, sheets);

			// excel代码创建
			string className = ExcelConverter.GetScriptClassString(excelName);
			string path = Path.ChangeExtension(Path.Combine(savePath, className), "cs");
			File.WriteAllText(path, scriptStrings[0]);

			// sheet代码创建
			for (int i = 0; i < sheets.Count; i++)
			{
				string entityClassName = GetSheetNameString(sheets[i].SheetName);

				// 检查类是否已存在
				string entityPath = Path.ChangeExtension(Path.Combine(savePath, entityClassName), "cs");
				if (!File.Exists(entityPath))
				{
					// 如果类不存在，创建并写入文件
					File.WriteAllText(entityPath, scriptStrings[i + 1]);
				}
				else
				{
					Debug.Log($"类 {entityClassName} 已存在，跳过创建。");
				}
			}

			// 刷新资源
			AssetDatabase.Refresh();
		}

		/// <summary>
		/// 创建ExcelAssetScript菜单的验证，判断是不是Excel表格，是的话才显示
		/// </summary>
		[MenuItem("Assets/Create/ExcelAssetScript", true)]
		static bool CreateScriptValidation()
		{
			var selectedAssets = Selection.GetFiltered(typeof(UnityEngine.Object), SelectionMode.Assets);
			if (selectedAssets.Length != 1) return false;
			var path = AssetDatabase.GetAssetPath(selectedAssets[0]);
			return Path.GetExtension(path) == ".xls" || Path.GetExtension(path) == ".xlsx";
		}

		/// <summary>
		/// 获取Excel表格的sheet名称
		/// </summary>
		/// <param name="excelPath">excel文件路径</param>
		/// <returns>所有的sheet名称组成的数组</returns>
		static List<ISheet> GetSheets(string excelPath)
		{
			var sheets = new List<ISheet>();
			using (FileStream stream = File.Open(excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
			{
				IWorkbook book = null;
				try
				{
					if (Path.GetExtension(excelPath).ToLower() == ".xls")
						book = new HSSFWorkbook(stream);
					else if (IsValidXlsxFile(excelPath))
						book = new XSSFWorkbook(stream);
					else
					{
						Debug.LogError($"Excel 文件格式不正确或损坏：{excelPath}");
						return new List<ISheet>();
					}
				}
				catch (Exception ex)
				{
					Debug.LogError($"读取 Excel 文件失败：{excelPath}\n{ex.GetType().Name}: {ex.Message}");
					throw;
				}

				for (int i = 0; i < book.NumberOfSheets; i++)
				{
					var sheet = book.GetSheetAt(i);
					sheets.Add(sheet);
				}
			}
			return sheets;
		}

		private static bool IsValidXlsxFile(string filePath)
		{
			try
			{
				using (FileStream fs = File.OpenRead(filePath))
				{
					byte[] header = new byte[2];
					int read = fs.Read(header, 0, 2);
					return read == 2 && header[0] == (byte)'P' && header[1] == (byte)'K';
				}
			}
			catch (Exception ex)
			{
				Debug.LogError($"检查文件头失败：{ex.Message}");
				return false;
			}
		}

		/// <summary>
		/// 获取表格代码模板字符串
		/// </summary>
		/// <returns>模板字符串</returns>
		static string GetScriptTempleteString()
		{
			string currentDirectory = Directory.GetCurrentDirectory();
			string[] filePath = Directory.GetFiles(currentDirectory, ScriptTemplateName, SearchOption.AllDirectories);
			if (filePath.Length == 0) throw new Exception("Script template not found.");

			string templateString = File.ReadAllText(filePath[0]);
			return templateString;
		}

		/// <summary>
		/// 获取Sheet代码模板字符串
		/// </summary>
		/// <returns>模板字符串</returns>
		static string GetEntityScriptTempleteString()
		{
			string currentDirectory = Directory.GetCurrentDirectory();
			string[] filePath = Directory.GetFiles(currentDirectory, SheetScriptTemplateName, SearchOption.AllDirectories);
			if (filePath.Length == 0) throw new Exception("Script template not found.");

			string templateString = File.ReadAllText(filePath[0]);
			return templateString;
		}



		/// <summary>
		/// 根据Excel表格名称和sheet名称创建对应的数据代码
		/// </summary>
		/// <param name="excelName">表格文件名</param>
		/// <param name="sheets">sheet名称</param>
		/// <returns>代码的字符串</returns>
		static List<string> BuildScriptStrings(string excelName, List<ISheet> sheets)
		{
			List<string> scriptStrings = new List<string>(sheets.Count + 1);

			string scriptString = GetScriptTempleteString();
			scriptStrings.Add(scriptString);

			string className = ExcelConverter.GetScriptClassString(excelName);
			scriptStrings[0] = scriptStrings[0].Replace("#ASSETSCRIPTNAME#", className);
			foreach (ISheet sheet in sheets)
			{
				string fieldString = String.Copy(FieldTemplete);
				fieldString = fieldString.Replace("#FIELDNAME#", sheet.SheetName);
				fieldString = fieldString.Replace("#EntityType#", GetSheetNameString(sheet.SheetName));

				fieldString += "\n#ENTITYFIELDS#";
				scriptStrings[0] = scriptStrings[0].Replace("#ENTITYFIELDS#", fieldString);

				string entityScriptString = BuildEntityScriptString(sheet);
				scriptStrings.Add(entityScriptString);
			}
			scriptStrings[0] = scriptStrings[0].Replace("#ENTITYFIELDS#\n", "");

			return scriptStrings;
		}

		/// <summary>
		/// 根据Sheet内容创建数据代码
		/// </summary>
		/// <param name="sheet">表格的Sheet</param>
		/// <returns>通过sheet生成的代码</returns>
		static string BuildEntityScriptString(ISheet sheet)
		{
			List<string> fieldNames = GetRowDataFromSheet(sheet, 0);
			List<string> fieldTypes = GetRowDataFromSheet(sheet, 1);

			string scriptString = GetEntityScriptTempleteString();

			scriptString = scriptString.Replace("#ASSETSCRIPTNAME#", GetSheetNameString(sheet.SheetName));

			StringBuilder variables = new StringBuilder(fieldNames.Count);
			for (int i = 0; i < fieldNames.Count; i++)
			{
				string variableName = fieldNames[i];
				string variableType = ConvertExcelTypeToCSharp(fieldTypes[i]);
				variables.Append($"\tpublic {variableType} {variableName};\n");
			}
			scriptString = scriptString.Replace("#ENTITYFIELDS#", variables.ToString());

			return scriptString;
		}

		/// <summary>
		/// 获取sheet对应的类名
		/// </summary>
		/// <param name="sheetName">sheet名称</param>
		/// <returns>类名</returns>
		private static string GetSheetNameString(string sheetName)
		{
			return string.Format(SheetScriptNameFormat, sheetName);
		}

		/// <summary>
		///	获取Excel表格的行数据
		/// </summary>
		/// <param name="row">行数</param>
		/// <returns>行内所有的值字符串数组</returns>
		private static List<string> GetRowDataFromSheet(ISheet sheet, int row)
		{
			IRow headerRow = sheet.GetRow(row);

			var fieldNames = new List<string>();
			for (int i = 0; i < headerRow.LastCellNum; i++)
			{
				var cell = headerRow.GetCell(i);
				if (cell == null || cell.CellType == CellType.Blank) break;
				fieldNames.Add(cell.StringCellValue);
			}
			return fieldNames;
		}


		/// <summary>
		/// 解析Entity字段的类型
		/// </summary>
		/// <param name="excelType">表格定义类型</param>
		/// <returns>代码数据类型</returns>
		private static string ConvertExcelTypeToCSharp(string excelType)
		{
			if (excelType.StartsWith("enum:", StringComparison.OrdinalIgnoreCase))
			{
				string enumName = excelType.Substring("enum:".Length);
				return enumName;
			}
			else
			{
				// 类型映射表
				Dictionary<string, string> typeMap = new Dictionary<string, string>()
				{
					// 基础数据类型
					{ "int", "int" },
					{ "float", "float" },
					{ "double", "double" },
					{ "string", "string" },
					{ "bool", "bool" },
					{ "char", "char" },
					{ "byte", "byte" },
					{ "short", "short" },
					{ "long", "long" },
					{ "decimal", "decimal" },

					// Unity 常用类型
					{ "vector2", "Vector2" },
					{ "vector3", "Vector3" },
					{ "vector4", "Vector4" },
				};
				return typeMap.TryGetValue(excelType.ToLower(), out string csharpType) ? csharpType : "string";
			}
		}
	}
}