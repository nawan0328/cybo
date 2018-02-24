package main

import (
	"fmt"
	"time"

	ole "github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

func main() {

	ole.CoInitialize(0)
	defer ole.CoUninitialize()
	unknown, _ := oleutil.CreateObject("CpUtil.CpCybos")
	CP_Session, _ := unknown.QueryInterface(ole.IID_IDispatch)

	conn, err := oleutil.GetProperty(CP_Session, "IsConnect")
	if err != nil {
		fmt.Println(err)
		return
	}

	iconn := i32toi(conn.Value())

	switch iconn {
	case 0 :
		fmt.Println("Disconnected")
		return
	case 1 :
		fmt.Println("Connected")
	}

	svrType, err := oleutil.GetProperty(CP_Session, "ServerType")
	if err != nil {
		fmt.Println(err)
		return
	}
	isvrType := i16toi(svrType.Value())

	switch isvrType {
	case 0 :
		fmt.Println("Disconnected")
		return
	case 1:
		fmt.Println("CyboPlus Server")
	case 2:
		fmt.Println("HTS Server")
	}

	checkPrice()

	/*for {
		fmt.Println("test")
		time.Sleep(1*time.Second)
	}*/
	time.Sleep(1*time.Second)
	fmt.Scanln()
}

func checkPrice(){

	//ole.CoInitialize(1)
	//defer ole.CoUninitialize()
	unknown, _ := oleutil.CreateObject("Dscbo1.StockMst")
	CP_StockPrice, _ := unknown.QueryInterface(ole.IID_IDispatch)

	_, err := oleutil.CallMethod(CP_StockPrice, "SetInputValue", 0, "A000660")
	if err != nil {
		fmt.Println(err)
		return
	}
	_, err = oleutil.CallMethod(CP_StockPrice,"BlockRequest")
	if err != nil {
		fmt.Println(err)
		return
	}
	code := oleutil.MustCallMethod(CP_StockPrice,"GetHeaderValue", 0)
	name := oleutil.MustCallMethod(CP_StockPrice,"GetHeaderValue", 1)
	price := oleutil.MustCallMethod(CP_StockPrice,"GetHeaderValue", 11)

	fmt.Println("Code : ", code.Value(), "Name : ", name.Value(), "Price : ", price.Value())

}

func i32toi(i32 interface{})(r int) {
	i32_tmp := i32.(int32)
	r = int(i32_tmp)
	return
}

func i16toi(i16 interface{})(r int) {
	i16_tmp := i16.(int16)
	r = int(i16_tmp)
	return
}
