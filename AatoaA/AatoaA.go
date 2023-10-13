package AatoaA

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"log"
	"strings"
)

func AatoaA() {
	// 打开 Excel 文件
	xlFile, err := excelize.OpenFile("AatoaA.xlsx")
	if err != nil {
		log.Fatal(err)
	}
	sheetList := xlFile.GetSheetList()
	if len(sheetList) == 0 {
		log.Fatal("Excel 文件中没有工作表")
	}

	// 获取第一个sheet表名
	sheetName := sheetList[0]
	//fmt.Println(sheetName)

	rows, err := xlFile.GetRows(sheetName)
	// 遍历每一行，处理第3列的值并保存到第6列
	for i, row := range rows {

		// 获取第3列的值
		value := row[2]

		// 转换大小写
		convertedValue := convertCase(value)

		// 保存到第6列
		xlFile.SetCellValue(sheetName, fmt.Sprintf("L%d", i+1), convertedValue)
	}

	// 保存修改后的Excel文件
	err = xlFile.Save()
	if err != nil {
		fmt.Println("保存Excel文件失败:", err)
		return
	}

	fmt.Println("大小写转换成功！")
}

func convertCase(str string) string {
	// 判断字符串中是否包含非英文字符
	hasNonEnglish := false
	for _, r := range str {
		if !strings.ContainsRune("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ", r) {
			hasNonEnglish = true
			break
		}
	}

	// 根据情况进行大小写转换
	if hasNonEnglish {
		return strings.Map(switchCase, str)
	}

	return str
}

func switchCase(r rune) rune {
	if r >= 'a' && r <= 'z' {
		return r - 32 // 转换为大写
	} else if r >= 'A' && r <= 'Z' {
		return r + 32 // 转换为小写
	}
	return r
}
