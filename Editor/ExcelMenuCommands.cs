using System.IO;
using UnityEditor;

namespace Le0der.Toolkits.Excel
{
    public static class ExcelMenuCommands
    {
        /// <summary>
        /// 根据excel表格创建对应的数据代码(顶部菜单)
        /// </summary>
        [MenuItem("Le0der Toolkits/Excel/生成数据代码", false, 100)]
        private static void CreateScriptFromMenu()
        {
            ExcelAssetScriptMenu.CreateScript();
        }

        /// <summary>
        /// 创建ExcelAssetScript菜单的验证，判断是不是Excel表格，是的话才显示
        /// </summary>
        [MenuItem("Le0der Toolkits/Excel/生成数据代码", true)]
        private static bool CreateScriptFromMenuValidation()
        {
            return CreateScriptValidation();
        }

        /// <summary>
        /// 重新导入选中的Excel表格
        /// </summary>
        [MenuItem("Le0der Toolkits/Excel/重新导入选中表格", false, 101)]
        public static void ReimportSelectedExcel()
        {
            var selectedAssets = Selection.GetFiltered(typeof(UnityEngine.Object), SelectionMode.Assets);
            if (selectedAssets.Length != 1) return;

            string path = AssetDatabase.GetAssetPath(selectedAssets[0]);
            string extension = Path.GetExtension(path).ToLower();

            if (extension == ".xls" || extension == ".xlsx")
            {
                // 执行重新导入操作，等同于右键 -> Reimport
                AssetDatabase.ImportAsset(path, ImportAssetOptions.ForceUpdate);
            }
        }

        /// <summary>
        /// 重新导入的菜单验证
        /// </summary>
        [MenuItem("Le0der Toolkits/Excel/重新导入选中表格", true)]
        private static bool ReimportSelectedValidation()
        {
            return CreateScriptValidation();
        }

        // /// <summary>
        // /// 根据excel表格创建对应的数据代码（右键菜单）
        // /// </summary>
        // [MenuItem("Assets/Create/ExcelAssetScript", false)]
        // static void CreateScript()
        // {
        //     ExcelAssetScriptMenu.CreateScript();
        // }


        // /// <summary>
        // /// 创建ExcelAssetScript菜单的验证，判断是不是Excel表格，是的话才显示
        // /// </summary>
        // [MenuItem("Assets/Create/ExcelAssetScript", true)]
        static bool CreateScriptValidation()
        {
            var selectedAssets = Selection.GetFiltered(typeof(UnityEngine.Object), SelectionMode.Assets);
            if (selectedAssets.Length != 1)
                return false;

            var path = AssetDatabase.GetAssetPath(selectedAssets[0]);
            string extension = Path.GetExtension(path).ToLower();
            return extension == ".xls" || extension == ".xlsx";
        }
    }
}