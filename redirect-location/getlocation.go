package redirect_location

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"io"
	"log"
	"net"
	"net/http"
)

func GetLocation() {

	// 打开 Excel 文件
	xlFile, err := excelize.OpenFile("test-casesentstive.xlsx")
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
		fmt.Println(i+1, URLA)

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
		fmt.Println("")

		// 获取 server Address 值，是个域名
		serverHost := resp.Request.URL.Host
		//fmt.Println("Remote Address:", remoteAddr)

		// 获取服务器IP地址
		serverIP, err := resolveIP(serverHost)
		if err != nil {
			fmt.Println("获取服务器IP地址出错：", err)
		}
		fmt.Println("服务器IP地址:", serverIP)
		fmt.Println("")

		xlFile.SetCellValue(sheetName, fmt.Sprintf("D%d", i+1), serverIP)
		xlFile.SetCellValue(sheetName, fmt.Sprintf("E%d", i+1), statusCode)
		xlFile.SetCellValue(sheetName, fmt.Sprintf("F%d", i+1), location)

	}
	// 保存修改后的 Excel 文件
	err = xlFile.Save()
	if err != nil {
		log.Fatal(err)
	}

}

// 解析域名获取IP地址
func resolveIP(hostname string) (string, error) {
	ips, err := net.LookupIP(hostname)
	if err != nil {
		return "", err
	}

	// 选择第一个非环回地址的IP
	for _, ip := range ips {
		if ip.To4() != nil && !ip.IsLoopback() {
			return ip.String(), nil
		}
	}

	return "", fmt.Errorf("无法解析IP地址")
}
