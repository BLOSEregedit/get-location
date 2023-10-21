package redirect_location

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"io"
	"log"
	"net"
	"net/http"
	"sync"
)

func GoGetLocation() {
	// 开始执行
	fmt.Println("")
	fmt.Println("*************************************")
	fmt.Println("")

	// 打开 Excel 文件
	xlFile, err := excelize.OpenFile("allre1016_prod.xlsx")
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

	// 创建一个 WatiGroup 来等待所有的 goroutine 完成
	var wg sync.WaitGroup

	/* 方式一
	var results chan Result
	results = make(chan Result, len(rows))
	*/

	results := make(chan Result, len(rows)) // results 是一个 channel,  len(rows) 是这个通道的缓冲大小
	mutex := sync.Mutex{}

	for i, row := range rows {
		wg.Add(1)
		go func(i int, row []string) {
			defer wg.Done()

			acheA := row[0]
			hostnameA := row[1]
			pathA := row[2]

			// 拼接 URLA ，打印行号和 URLA
			URLA := acheA + hostnameA + pathA
			//fmt.Println(i+1, URLA)

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
			//defer resp.Body.Close() // 添加这一行来关闭响应的 Body
			defer func(Body io.ReadCloser) {
				err := Body.Close()
				if err != nil {
					log.Println(err)
				}
			}(resp.Body)

			// 获取响应状态码
			statusCode := resp.StatusCode
			//fmt.Println("状态码:", statusCode)

			// 获取 Location 值
			location := resp.Header.Get("Location")
			//fmt.Println("Location:", location)
			//fmt.Println("")

			// 获取 server Address 值，是个域名
			serverHost := resp.Request.URL.Host
			//fmt.Println("Remote Address:", remoteAddr)

			// 获取服务器IP地址
			serverIP, err := goresolveIP(serverHost)
			if err != nil {
				fmt.Println("获取服务器IP地址出错：", err)
			}
			//fmt.Println("服务器IP地址:", serverIP)
			//fmt.Println("")

			result := Result{ // 此为一个结构体 Result，将刚刚单独获取到的变量值进行了初始化，并进行相应赋值
				Index:      i,
				ServerIP:   serverIP,
				StatusCode: statusCode,
				Location:   location,
			}

			mutex.Lock()
			results <- result // result 的值发送给了 results 这个 channel
			mutex.Unlock()

			fmt.Println(i+1, URLA)
			fmt.Println("状态码:", statusCode)
			fmt.Println("Location:", location)
			fmt.Println("服务器IP地址:", serverIP)
			fmt.Println("")

		}(i, row)
	}

	go func() {
		wg.Wait()
		close(results)
	}()

	for result := range results {
		mutex.Lock()
		xlFile.SetCellValue(sheetName, fmt.Sprintf("D%d", result.Index+1), result.ServerIP)
		xlFile.SetCellValue(sheetName, fmt.Sprintf("E%d", result.Index+1), result.StatusCode)
		xlFile.SetCellValue(sheetName, fmt.Sprintf("F%d", result.Index+1), result.Location)
		mutex.Unlock()
	}

	// 保存修改后的 Excel 文件
	err = xlFile.Save()
	if err != nil {
		log.Fatal(err)
	}
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

// 结果结构体
type Result struct {
	Index      int
	ServerIP   string
	StatusCode int
	Location   string
}
