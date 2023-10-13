package main

import (
	"get_location/ascii"
	"get_location/compare"
	redirect_location "get_location/redirect-location"
)

func main() {
	//AatoaA.AatoaA()

	//queryescape.Queryescape()

	//redirect_location.GetLocation()
	////
	//ascii.DoAscii()

	redirect_location.GoGetLocation() // 并发，获取中间状态 状态码 + location
	compare.GoCompare()               // 比对单元格值是否一致

	ascii.GoDoAscii() // 并发，获取最终请求的状态码

}
