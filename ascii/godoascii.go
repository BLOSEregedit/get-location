package ascii

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"log"
	"net/http"
)

func GoDoAscii() {

	// 打开 Excel 文件
	xlFile, err := excelize.OpenFile("prod-eu-page2-无空白.xlsx")
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
		resp, err := http.Get(URLA)
		if err != nil {
			log.Fatal(err)
		}

		// 获取状态码
		status := resp.StatusCode
		fmt.Println("Status Code:", status)

		// 将请求 URL 保存到变量 ascii
		ascii := resp.Request.URL.String()
		fmt.Println("Request URL:", ascii)

		// 关闭响应体
		resp.Body.Close()
		fmt.Println("")
		fmt.Println("")

		// 在当前工作表中保存状态码和请求 URL
		xlFile.SetCellValue(sheetName, fmt.Sprintf("D%d", i+1), status)
		xlFile.SetCellValue(sheetName, fmt.Sprintf("E%d", i+1), ascii)

		// 保存修改后的 Excel 文件
		err = xlFile.Save()
		if err != nil {
			log.Fatal(err)
		}

	}

	//// 保存修改后的 Excel 文件
	//err = xlFile.Save()
	//if err != nil {
	//	log.Fatal(err)
	//}

}
