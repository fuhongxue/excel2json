package main

import (
	"encoding/json"
	"errors"
	"fmt"
	"github.com/bitly/simplejson"
	"github.com/tealeg/xlsx"
	"io/ioutil"
	"log"
	"os"
	"path/filepath"
	"regexp"
	"strings"
)

//配置结构
type Excel2Json struct {
	curPath     string           //当前路径
	jsonDir     string           //json存放目录
	excelDir    string           //excel源目录
	showWarning bool             //转换时是否显示警告
	jsonConf    *simplejson.Json //配置读取
	suc         []baseMsg        //转换成功记录
	fail        []baseMsg        //转换失败记录
	warn        []baseMsg        //配置可能有误的记录
}

//记录转换成功和失败内容
type baseMsg struct {
	fname string //原文件名
	oname string //生成文件名
	err   error  //转换结果提示语
}

var cfg *Excel2Json

func NewExcel2Json() *Excel2Json {
	path, err := os.Getwd() //当前路径
	if err != nil {
		log.Fatalln("获取路出错.")
	}
	f, err := os.Open(filepath.Join(path, "conf.json")) //JSON配置文件
	if err != nil {
		log.Panicln("读取配置文件失败.")
	}
	defer f.Close()
	fd, err := ioutil.ReadAll(f)
	jsonConf, err := simplejson.NewJson(fd)
	if err != nil {
		log.Panicln("解析配置文件出错.")
	}
	excel_dir, _ := jsonConf.Get("excel_dir").String()
	excel_dir, _ = filepath.Abs(excel_dir)
	json_dir, _ := jsonConf.Get("json_dir").String()
	json_dir, _ = filepath.Abs(json_dir)
	os.MkdirAll(json_dir, 0666) //创建json存放目录
	show_warning, _ := jsonConf.Get("show_warning").Bool()

	return &Excel2Json{
		curPath:     path,               //当前路径
		jsonDir:     json_dir,           //json存放目录
		excelDir:    excel_dir,          //excel源目录
		showWarning: show_warning,       //转换时是否显示警告
		jsonConf:    jsonConf,           //配置读取
		suc:         make([]baseMsg, 0), //转换成功记录
		fail:        make([]baseMsg, 0), //转换成功失败
		warn:        make([]baseMsg, 0), //配置可能有误的记录
	}
}

func NewExcelData() *excelData {
	return &excelData{
		headMap:       make(map[int]string, 0),
		data:          make([]map[string]interface{}, 0),
		hasDesc:       true,
		hasHead:       false,
		descCount:     0,
		headLineIndex: 0,
	}
}

type excelData struct {
	headMap       map[int]string           //表头 map[下标]表头字符串
	data          []map[string]interface{} //数据 []map[表头字符串]数据
	hasDesc       bool                     //前几行是否还有描述,默认为有
	hasHead       bool                     //是否有表头,默认为无
	descCount     int                      //excel备注有几行
	headLineIndex int                      //表头在excel第几行
}

func (e *Excel2Json) excel2json() {
	defer func() {
		if x := recover(); x != nil {
			log.Printf("【致命错误|必需解决】:%s\n", x)
		}
	}()
	filepath.Walk(e.excelDir, func(fname string, info os.FileInfo, err error) error {
		if !info.IsDir() {
			fstr := info.Name()
			if strings.HasSuffix(fstr, ".xlsx") && string(fstr[0]) != "~" {
				index := strings.LastIndex(fstr, ".")
				fn := fstr[:index]
				key := fn[strings.LastIndexAny(fn, "-")+1:] //原文件名

				baseFile := filepath.Base(fstr)
				outBaseFile := key + ".json" //生成文件名
				log.Printf("(读取) 源文件:【%s】\n", baseFile)
				f, err := xlsx.OpenFile(fname)
				if err != nil {
					e.fail = append(e.fail, baseMsg{fname: baseFile, oname: outBaseFile, err: err})
				}
				for _, sheet := range f.Sheets {

					//////////////////////////////////////////

					myExp := regexp.MustCompile("[\u4e00-\u9fa5]+") //中文正则
					eData := NewExcelData()
					cols := 0
					for i, row := range sheet.Rows {
						cols = len(row.Cells)                      //总列数
						tmpData := make(map[string]interface{}, 0) //每行的临时数组
						//得到某行各列数据
						for j, cell := range row.Cells {
							//过滤特殊字符
							s := strings.Trim(strings.Replace(cell.String(), "\n", "", -1), "|")
							if j == 0 && s == "" { //跳过某行第一列为空的行
								break
							}
							if eData.hasDesc && myExp.MatchString(s) { //过滤前几行的备注
								eData.descCount++
								break
							} else {
								eData.hasDesc = false
							}
							if !eData.hasHead { //过滤空行和备注的第一行必为表头
								eData.headLineIndex = i + 1 //原文件excel第几行
								eData.headMap[j] = s
							} else { //内容
								if eData.headMap[j] == "" {
									e.warn = append(e.warn, baseMsg{fname: baseFile, oname: outBaseFile, err: errors.New(fmt.Sprintf("(%d行,%d列)表头为空?【单元格内容:%s】", i+1, j+1, s))})
								} else {
									tmpData[eData.headMap[j]] = s
								}
							}
						}

						if eData.hasHead { //数据
							eData.data = append(eData.data, tmpData)
						} else { //表头
							hlen := len(eData.headMap)
							if hlen > 0 { //已经有填充表头
								eData.hasHead = true
								if len(row.Cells) != len(eData.headMap) { //表头和列数不相等,可能配置有误
									e.warn = append(e.warn, baseMsg{fname: baseFile, oname: outBaseFile, err: errors.New(fmt.Sprintf("表头数目(%d)小于列数目(%d)", hlen, cols))})
								}
							} else { //一些备注和空行

							}
						}
					}
					b, _ := json.Marshal(eData.data)
					b = []byte("{\"Data\":" + string(b) + "}")
					err = ioutil.WriteFile(filepath.Join(e.jsonDir, outBaseFile), b, 0666)
					if err != nil {
						e.fail = append(e.fail, baseMsg{fname: baseFile, oname: outBaseFile, err: err})
					}
					/////////////////////////////////////////
					break //只读sheet1
				}

				e.suc = append(e.suc, baseMsg{fname: baseFile, oname: outBaseFile, err: nil})
			}
		}
		return nil
	})

}

func (e *Excel2Json) PrintAllMsg() {
	slen, flen, wlen := len(e.suc), len(e.fail), len(e.warn)
	log.Printf("(检测) excel源目录:【%s】\n", e.excelDir)
	log.Printf("(检测) json生成目录:【%s】\n", e.jsonDir)
	log.Printf("(检测) 一共扫描到【%d】个excel文件.\n\n", slen+flen)
	if slen > 0 {
		log.Printf("(成功) 转换=======【%d】个excel文件成功===========:\n", slen)
		for _, v := range e.suc {
			log.Printf("(成功) 转换:【%s】=>【%s】\n", v.fname, v.oname)
		}
	}

	if flen > 0 {
		log.Printf("(失败) 转换=======【%d】个excel文件失败===========:\n", flen)
		for _, v := range e.fail {
			log.Printf("(失败) 【%s】=>【%s】,出错:%s\n", v.fname, v.oname, v.err.Error())
		}
	}
	if e.showWarning {
		if wlen > 0 {
			log.Printf("(警告) 转换产生=======【%d】个警告,某些excel配置可能有误=======:\n", wlen)
			for _, v := range e.warn {
				log.Printf("(警告) 【%s】,提示:%s\n", v.fname, v.err.Error())
			}
		}
	} else {
		log.Printf("警告关闭,打开警告请将show_warning设为true.\n\n")
	}
	log.Println("Done.\n\n\t\t\t【Excel转json工具-蓝条V1.0】\n\n\t\t\t【回车退出】")
}

func main() {
	cfg = NewExcel2Json()
	cfg.excel2json()
	cfg.PrintAllMsg()

	b := make([]byte, 1)
	os.Stdin.Read(b)
}
