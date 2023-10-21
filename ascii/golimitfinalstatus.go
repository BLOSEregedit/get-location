package ascii

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"log"
	"net"
	"net/http"
	"sync"
	"time"
)

var (
	maxConcurrency = 400 // 并发限制数
	client         *http.Client
	connPool       *sync.Pool
)

func init() {
	// 创建具有连接池的 HTTP 客户端
	client = &http.Client{
		Transport: &http.Transport{
			MaxIdleConns:        maxConcurrency,
			MaxIdleConnsPerHost: maxConcurrency,
		},
		Timeout: time.Second * 30, // 请求超时时间
	}

	// 创建连接池
	connPool = &sync.Pool{
		New: func() interface{} {
			conn, err := excelize.OpenFile("allre1016_prod.xlsx")
			if err != nil {
				log.Fatal(err)
			}
			return conn
		},
	}
}

func GoLimitFinalstatus() {
	fmt.Println("")
	fmt.Println("*************************************")
	fmt.Println("")

	// 从连接池获取 Excel 连接
	conn := connPool.Get().(*excelize.File)
	defer connPool.Put(conn) // 将连接放回连接池

	sheetList := conn.GetSheetList()
	if len(sheetList) == 0 {
		log.Fatal("Excel 文件中没有工作表")
	}

	// 获取第一个sheet表名
	sheetName := sheetList[0]

	// 读取 sheetName 表
	rows, err := conn.GetRows(sheetName)
	if err != nil {
		log.Fatal(err)
	}

	var wg sync.WaitGroup
	mutex := sync.Mutex{}
	semaphore := make(chan struct{}, maxConcurrency) // 控制并发数的信号量

	for i, row := range rows {
		wg.Add(1)
		semaphore <- struct{}{} // 获取信号量

		go func(i int, row []string) {
			defer func() {
				<-semaphore // 释放信号量
				wg.Done()
			}()

			acheA := row[0]
			hostnameA := row[1]
			pathA := row[2]

			// 拼接 URLA
			URLA := acheA + hostnameA + pathA

			// 发送 GET 请求
			status, serverHost := golimitGetRequest(URLA)

			// 获取服务器IP地址
			serverIP, err := golimitresolveIP(serverHost)
			if err != nil {
				fmt.Println("获取服务器IP地址出错：", err)
			}
			fmt.Println("服务器IP地址:", serverIP)
			fmt.Println("")

			mutex.Lock()
			// 在当前工作表中保存状态码和请求 URL
			conn.SetCellValue(sheetName, fmt.Sprintf("H%d", i+1), serverIP)
			conn.SetCellValue(sheetName, fmt.Sprintf("I%d", i+1), status)
			mutex.Unlock()

			// 集中打印
			fmt.Println(i, URLA)
			fmt.Println("服务器IP地址:", serverIP)
			fmt.Println("Status Code:", status)

			fmt.Println("")
		}(i, row)
	}

	wg.Wait()

	// 异步保存修改后的 Excel 文件
	go func() {
		err := conn.Save()
		if err != nil {
			log.Fatal(err)
		}
	}()

	// 等待保存操作完成
	time.Sleep(time.Second)

	fmt.Println("Excel 文件保存成功！")
}

func golimitGetRequest(URLA string) (int, string) {
	defer func() {
		if r := recover(); r != nil {
			log.Println("发生错误:", r)
		}
	}()

	resp, err := client.Get(URLA)
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
