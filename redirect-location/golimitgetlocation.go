package redirect_location

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"log"
	"net"
	"net/http"
	"sync"
	"time"
)

// 最大并发数
const maxConcurrency = 400

// HTTP连接池
var httpClient *http.Client

func init() {
	// 创建自定义的HTTP客户端
	httpClient = &http.Client{
		Transport: &http.Transport{
			MaxIdleConns:        maxConcurrency, // 设置最大空闲连接数
			MaxIdleConnsPerHost: maxConcurrency, // 设置每个主机的最大空闲连接数
		},
		Timeout: time.Second * 60, // 设置超时时间 X 秒
	}
}

func GoLimitGetLocation() {
	// 开始执行
	fmt.Println("")
	fmt.Println("*************************************")
	fmt.Println("")

	// 打开 Excel 文件
	xlFile, err := excelize.OpenFile("allre1016_double.xlsx")
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
	results := make(chan Result, len(rows))
	mutex := sync.Mutex{}
	semaphore := make(chan struct{}, maxConcurrency) // 控制并发数的信号量

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
			req, err := http.NewRequest("GET", URLA, nil)
			if err != nil {
				fmt.Println("创建请求出错：", err)
				return
			}
			req.Header.Set("Connection", "close") // 关闭连接以避免连接保持问题

			// 控制并发数，获取信号量
			semaphore <- struct{}{}
			defer func() {
				// 释放信号量
				<-semaphore
			}()

			// 发送请求
			resp, err := httpClient.Do(req)
			if err != nil {
				fmt.Println("请求出错：", err)
				return
			}
			defer func() {
				err := resp.Body.Close()
				if err != nil {
					log.Println(err)
				}
			}()

			// 获取响应状态码
			statusCode := resp.StatusCode

			// 获取 Location 值
			location := resp.Header.Get("Location")

			// 获取服务器IP地址
			serverIP, err := golimitresolveIP(hostnameA)
			if err != nil {
				fmt.Println("获取服务器IP地址出错：", err)
			}

			result := Result{
				Index:      i,
				ServerIP:   serverIP,
				StatusCode: statusCode,
				Location:   location,
			}

			mutex.Lock()
			results <- result
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
func golimitresolveIP(hostname string) (string, error) {
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
type ResultLimit struct {
	Index      int
	ServerIP   string
	StatusCode int
	Location   string
}
