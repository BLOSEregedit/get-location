package ascii

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"log"
	"net"
	"net/http"
	"sync"
)

func GoDoAscii() {
	// 打开 Excel 文件
	xlFile, err := excelize.OpenFile("gogetlocation.xlsx")
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

	var wg sync.WaitGroup
	mutex := sync.Mutex{}

	for i, row := range rows {
		wg.Add(1)
		go func(i int, row []string) {
			defer wg.Done()

			acheA := row[0]
			hostnameA := row[1]
			pathA := row[2]

			// 拼接 URLA
			URLA := acheA + hostnameA + pathA
			//fmt.Println(i, URLA)

			// 发送 GET 请求
			status, serverHost := gosendGetRequest(URLA)

			// 获取服务器IP地址
			serverIP, err := goresolveIP(serverHost)
			if err != nil {
				fmt.Println("获取服务器IP地址出错：", err)
			}
			fmt.Println("服务器IP地址:", serverIP)
			fmt.Println("")

			mutex.Lock()
			// 在当前工作表中保存状态码和请求 URL
			xlFile.SetCellValue(sheetName, fmt.Sprintf("H%d", i+1), serverIP)
			xlFile.SetCellValue(sheetName, fmt.Sprintf("I%d", i+1), status)
			mutex.Unlock()

			// 集中打印
			fmt.Println(i, URLA)
			fmt.Println("服务器IP地址:", serverIP)
			fmt.Println("Status Code:", status)

			fmt.Println("")
		}(i, row)
	}

	wg.Wait()

	// 保存修改后的 Excel 文件
	err = xlFile.Save()
	if err != nil {
		log.Fatal(err)
	}
}

func gosendGetRequest(URLA string) (int, string) {
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
	serverHost := resp.Request.URL.Host

	resp.Body.Close()
	return status, serverHost
}

// 解析域名获取IP地址
func goresolveIP(hostname string) (string, error) {
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
