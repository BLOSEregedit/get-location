package main

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"io"
	"log"
	"net/http"
)

func main() {

	// 打开 Excel 文件
	xlFile, err := excelize.OpenFile("prod-eu-page-错误url.xlsx")
	if err != nil {
		log.Fatal(err)
	}
	sheetList := xlFile.GetSheetList()
	if len(sheetList) == 0 {
		log.Fatal("Excel 文件中没有工作表")
	}

	// 获取第一个sheet表名
	sheetName := sheetList[0]

	// 创建新的工作表以保存结果
	newSheetName := "result"
	_, err = xlFile.NewSheet(newSheetName)
	if err != nil {
		log.Fatal(err)
	}

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

		// 发送 GET 请求到 URLA，禁止自动重定向
		client := &http.Client{
			CheckRedirect: func(req *http.Request, via []*http.Request) error {
				return http.ErrUseLastResponse
			},
		}

		// 发送 GET 请求到 URLA
		resp, err := client.Get(URLA)
		if err != nil {
			fmt.Println("请求出错：", err)
			return
		}
		defer func(Body io.ReadCloser) {
			err := Body.Close()
			if err != nil {
				log.Println(err)
			}
		}(resp.Body)

		// 获取响应状态码
		statusCode := resp.StatusCode
		fmt.Println("状态码:", statusCode)

		// 获取 Location 值
		location := resp.Header.Get("Location")
		fmt.Println("Location:", location)

	}

}
