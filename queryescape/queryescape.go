package queryescape

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"html"
	"log"
	"net/url"
	"strings"
)

func Queryescape() {
	// 打开 Excel 文件
	f, err := excelize.OpenFile("prod-eu-page2_1011rd200.xlsx")
	if err != nil {
		log.Fatal(err)
	}

	sheetList := f.GetSheetList()
	if len(sheetList) == 0 {
		log.Fatal("Excel 文件中没有工作表")
	}

	// 获取第一个sheet表名
	sheetName := sheetList[0]

	// 读取工作表中的 C 列数据
	cols, err := f.GetCols(sheetName)
	if err != nil {
		log.Fatal(err)
	}

	// 遍历 C 列数据
	for rowIndex, cellValue := range cols[2] {
		// 进行 URL 转义，排除对 / 字符的转义
		escaped := customEscape(cellValue)

		// 进行 HTML 转义
		escaped = html.EscapeString(escaped)

		// 将转义后的结果保存在 D 列
		f.SetCellValue(sheetName, fmt.Sprintf("D%d", rowIndex+1), escaped)
	}

	// 保存修改后的 Excel 文件
	err = f.Save()
	if err != nil {
		log.Fatal(err)
	}

	fmt.Println("转义完成")
}

// 自定义的 URL 转义函数，忽略 /
func customEscape(s string) string {
	// 首先使用 url.QueryEscape 进行转义
	escaped := url.PathEscape(s)

	// 将转义后的结果中的 %2F 替换回 /
	escaped = strings.ReplaceAll(escaped, "%2F", "/")

	return escaped
}
