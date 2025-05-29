using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityEngine.UI;

public class Example : MonoBehaviour
{
	[SerializeField] ExcelMstItems mstItems;
	[SerializeField] Text text;

	void Start()
	{
		ShowItems();
	}

	void ShowItems()
	{
		string str = "";

		mstItems.Entity
			.ForEach(entity => str += DescribeMstItemEntity(entity) + "\n");

		text.text = str;
	}

	string DescribeMstItemEntity(SheetEntityEntity entity)
	{
		return string.Format(
			"{0} : {1}, {2}, {3}, {4}, {5}",
			entity.id,
			entity.name,
			entity.price,
			entity.isNotForSale,
			entity.rate,
			entity.category
		);
	}
}

public enum MstItemCategory
{
	Red,
	Green,
	Blue,
}

