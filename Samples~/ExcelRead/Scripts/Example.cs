using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityEngine.UI;

namespace Le0der.Toolkits.Excel.Demo
{
	public class Example : MonoBehaviour
	{
		[SerializeField] ExcelSample sample;
		[SerializeField] Text text;

		void Start()
		{
			ShowItems();
		}

		void ShowItems()
		{
			string str = "";

			sample.Sample
				.ForEach(sample => str += DescribeMstItemEntity(sample) + "\n");

			text.text = str;
		}

		string DescribeMstItemEntity(SheetEntitySample sample)
		{
			return string.Format(
				"{0} : {1}, {2}, {3}, {4}, {5}",
				sample.id,
				sample.name,
				sample.price,
				sample.isNotForSale,
				sample.rate,
				sample.category
			);
		}
	}

	public enum MstItemCategory
	{
		Red,
		Green,
		Blue,
	}
}