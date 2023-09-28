package ascii

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"log"
	"net/http"
)

func DoAscii() {
	// 打开 Excel 文件
	xlFile, err := excelize.OpenFile("prod-eu-page2-处理后.xlsx")
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

	for i, row := range rows {
		acheA := row[0]
		hostnameA := row[1]
		pathA := row[2]

		// 拼接 URLA
		URLA := acheA + hostnameA + pathA
		fmt.Println(i, URLA)

		// 发送 GET 请求
		status, ascii := sendGetRequest(URLA)

		// 在当前工作表中保存状态码和请求 URL
		xlFile.SetCellValue(sheetName, fmt.Sprintf("D%d", i+1), status)
		xlFile.SetCellValue(sheetName, fmt.Sprintf("E%d", i+1), ascii)

		fmt.Println("Status Code:", status)
		fmt.Println("Request URL:", ascii)
		fmt.Println("")
		fmt.Println("")
	}

	// 保存修改后的 Excel 文件
	err = xlFile.Save()
	if err != nil {
		log.Fatal(err)
	}
}

func sendGetRequest(URLA string) (int, string) {
	defer func() {
		if r := recover(); r != nil {
			log.Println("发生错误:", r)
		}
	}()

	resp, err := http.Get(URLA)
	if err != nil {
		log.Println("请求出错:", err)
		return 0, ""
	}

	status := resp.StatusCode
	ascii := resp.Request.URL.String()

	resp.Body.Close()
	return status, ascii
}
