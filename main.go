package main

import (
	//"database/sql"
	//	"encoding/json"
	"fmt"
	//"io/ioutil"
	"strconv"
	"time"

	//"github.com/astaxie/beego"
	"github.com/astaxie/beego/orm"

	_ "github.com/go-sql-driver/mysql"
)

func main() {
	start := time.Now()

	maxIdle := 1
	maxConn := 4

	orm.RegisterDataBase("default", "mysql", "root:mypassword@tcp(10.0.12.104:3306)/sale_stat?loc=Local", maxIdle, maxConn)
	//orm.RunSyncdb("default", false, true)
	orm.Debug = true

	//resultPointer, columnsPointer := sqlFetch(db, query)

	values, err := ExcelSql("select * from temp_model limit 0,10")

	var column_name_map map[string]string
	// 再使用make函数创建一个非nil的map，nil map不能赋值
	column_name_map = make(map[string]string)
	column_name_map["chart_type"] = "图形类型"

	if err == nil && len(values) > 0 {
		ExportExcel(&values, "abc.xlsx", column_name_map)
	}

	//excel(resultPointer, columnsPointer)
	end := time.Now()
	fmt.Println("total time : ", timeFriendly(end.Sub(start).Seconds()))
}

// time format fridnely
func timeFriendly(second float64) string {

	if second < 1 {
		return strconv.Itoa(int(second*1000)) + "毫秒"
	} else if second < 60 {
		return strconv.Itoa(int(second)) + "秒" + timeFriendly(second-float64(int(second)))
	} else if second >= 60 && second < 3600 {
		return strconv.Itoa(int(second/60)) + "分" + timeFriendly(second-float64(int(second/60)*60))
	} else if second >= 3600 && second < 3600*24 {
		return strconv.Itoa(int(second/3600)) + "小时" + timeFriendly(second-float64(int(second/3600)*3600))
	} else if second > 3600*24 {
		return strconv.Itoa(int(second/(3600*24))) + "天" + timeFriendly(second-float64(int(second/(3600*24))*(3600*24)))
	}
	return ""
}

func ExcelSql(sql string) ([]orm.Params, error) {
	var values []orm.Params
	var num int64
	var err error
	o := orm.NewOrm()

	num, err = o.Raw(sql).Values(&values)

	if num > 0 && err == nil {
		return values, nil
	}
	return nil, err
}
