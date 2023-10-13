package compare

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"log"
)

func GoCompare() {
	// 打开 Excel 文件
	xlFile, err := excelize.OpenFile("compare1.xlsx")
	if err != nil {
		log.Fatal(err)
	}
	sheetList := xlFile.GetSheetList()
	if len(sheetList) == 0 {
		log.Fatal("Excel 文件中没有工作表")
	}

	// 获取第一个sheet表名
	sheetName := sheetList[0]

	// 读取 sheetName 表
	rows, err := xlFile.GetRows(sheetName)
	if err != nil {
		log.Fatal(err)
	}

	// 遍历每一行进行比较
	for i, row := range rows {
		// 获取 F 列和 K 列的值
		fValue := row[5]  // F 列索引为 5
		kValue := row[10] // K 列索引为 10

		// 比较 F 列和 K 列的值
		var result int
		if fValue == kValue {
			result = 1
		} else {
			result = 0
		}

		// 在对应行的 L 列记录比较结果
		cell := fmt.Sprintf("L%d", i+1) // i+1 表示行号，从第一行开始
		xlFile.SetCellValue(sheetName, cell, result)

		fmt.Printf("行数:%4d   结果：%d\n", i+1, result)
	}

	// 保存修改后的 Excel 文件
	err = xlFile.Save()
	if err != nil {
		log.Fatal(err)
	}
}
