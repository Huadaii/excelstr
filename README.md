# Excelstr
将Go Struct 转换为Excel的工具  
Tools to convert Go Struct to Excel  

使用教程  
Method of use  
  
需要给结构体加上`xlsx:"name"` Tag  
You need to add `xlsx:"name"` Tag to the structure  
  
```  
type NetworkInfo struct {  
	NetCards       []NetCard        `json:"网卡状况" xlsx:"网卡信息"` //网卡状况  
	ConnectionStat []ConnectionStat `json:"netstat" xlsx:"连接信息"`  
}
  
type NetCard struct {  
	Device  int    `xlsx:"网卡设备"` //网卡设备  
	Name    string `xlsx:"网卡名称"`  
	Type    string `xlsx:"网络类型"`     //网络类型  
	UUID    string `xlsx:"标识符"`      //标识符  
	IPAddr  string `xlsx:"IP地址"`     //ip地址  
	HWAddr  string `xlsx:"Mac地址"`    //Mac地址  
	NetMask string `xlsx:"子网掩码"`     //子网掩码  
	GateWay string `xlsx:"网关"`       //网关  
	IPType  string `xlsx:"网卡获取ip方式"` //网卡获取ip方式  dhcp ,none ,static  
	OnBoote bool   `xlsx:"开机自启动"`    //开机自启动  
}  
````  
  
## 调用方法  
```
//Excel 一层结构体
func Excel(str interface{}, sheetName string) *excelize.File {}
//ExcelStruct 嵌套结构体
func ExcelStruct(str interface{}) *excelize.File {}  

```  
  
![image](https://user-images.githubusercontent.com/51690238/165208394-07128aed-8308-49b5-b7f6-593910e5535e.png)
