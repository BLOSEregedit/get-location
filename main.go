package main

import (
	"get_location/ascii"
	"get_location/compare"
	redirect_location "get_location/redirect-location"
)

func main() {
	redirect_location.GetLocation()
	//queryescape.Queryescape()

	//redirect_location.GoLimitGetLocation()

	//redirect_location.GoGetLocation() // 并发，获取中间状态 状态码 + location
	compare.GoCompare() // 比对单元格值是否一致
	//compare.GoCompare()

	////ascii.DoAscii()
	//ascii.GoDoAscii() // 并发，获取最终请求的状态码

	ascii.GoLimitFinalstatus()

}
