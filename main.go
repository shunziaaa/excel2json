package main

import (
	"encoding/json"
	"fmt"
	"github.com/Luxurioust/excelize"
	"io"
	"io/ioutil"
	"log"
	"os"
	"path"
	"strings"
	_ "strings"
)

type ConfigStu struct {
	ReadDir string `json: "ReadDir"`
	ReadSuffix string `json: ReadSuffix`
	BuildDir string `json: BuildDir`
	BuildSuffix string `json: BuildSuffix`
}
var Config ConfigStu
func initLog() int {
	fileName := "./Log.txt"
	fileHandle,err := os.OpenFile(fileName, os.O_CREATE|os.O_WRONLY|os.O_TRUNC, 0666)
	if err != nil {
		fmt.Println("Error: initLog ", err)
		return -1
	}
	wIo := io.MultiWriter(fileHandle, os.Stdout)
	log.SetOutput(wIo)
	log.SetFlags(log.LstdFlags|log.Lshortfile)
	return 0
}

func getConfig() int {
	fData, err := ioutil.ReadFile("./config.json")
	if err != nil{
		log.Printf("get config error: %s\n", err)
		return -1
	}

	err = json.Unmarshal(fData, &Config)
	if err != nil {
		log.Printf("config data init error : %s\n", err)
		return -1
	}
	return 0
}

func excel2Json (fileName string) int {
	buildHandle, err := os.Stat(Config.BuildDir)
	if err != nil || !buildHandle.IsDir()  {
		os.Mkdir(Config.BuildDir, os.ModePerm)
	}
	fileAllName := Config.ReadDir + "/" + fileName
	excelFile, err := excelize.OpenFile(fileAllName)
	if err != nil {
		log.Printf("read %s error: %s\n", fileName, err)
		return -1
	}

	Rows, err := excelFile.GetRows("Sheet1")
	if len(Rows) <= 1 {
		log.Printf("%s is null\n", fileName)
		return -1
	}

	typeList := make([]string, 0)
	nameList := make([]string, 0)
	ignoreMap := make(map[int]int, 0)

	for _,v := range Rows[0]{
		typeList = append(typeList, v)
	}

	for index,v := range Rows[1]{
		nameList = append(nameList, v)
		if v[0] == '#' {
			ignoreMap[index] = 0
		}
	}

	resStr := "{"
	rowsLen := len(Rows)
	rowLen := len(Rows[0])

	for true {
		_, isExist := ignoreMap[rowLen -1]
		if !isExist {
			break
		}
		rowLen --
	}

	for index, Row := range Rows{
		if index == 0 || index == 1{
			continue
		}
		for k, data := range Row {
			_, exist := ignoreMap[k]
			if exist {
				continue
			}
			if k == 0{
				resStr += "\""
				if data == "" {
					log.Printf("%s key is null, index: %d\n", fileName, k)
					return -1
				}
				resStr += data
				resStr += "\""
				resStr += ":{"
			}

			if typeList[k] == "int"{
				resStr += "\""
				resStr += nameList[k]
				resStr += "\""
				resStr += ":"
				if data == "" {
					data = "0"
				}
				resStr += data

			} else if typeList[k] == "string" {
				resStr += "\""
				resStr += nameList[k]
				resStr += "\""
				resStr += ":"
				resStr += "\""
				resStr += data
				if data == "" {
					data = "\"\""
				}
				resStr += "\""
			}else{
				log.Printf("%s not type %s\n", fileName, typeList[k])
				return -1
			}
			if k != rowLen -1 {
				resStr += ","
			}
		}
		resStr += "}"
		if index != rowsLen - 1 {
			resStr += ","
		}
	}
	resStr += "}"

	buildName :=Config.BuildDir + "/" + strings.TrimSuffix(fileName, Config.ReadSuffix) + Config.BuildSuffix
	openFile,err := os.OpenFile(buildName, os.O_WRONLY|os.O_CREATE|os.O_TRUNC, 0666)
	if err != nil {
		log.Printf("%s create error: %s", fileName, err)
		return -1
	}
	defer openFile.Close()
	openFile.Write([]byte(resStr))

	return 0
}

func run() int {
	files, err := ioutil.ReadDir(Config.ReadDir)
	if err != nil {
		log.Printf("Error read dir error: %s\n", err)
		return -1
	}

	for _, file := range files{
		if path.Ext(file.Name()) != Config.ReadSuffix{
			continue
		}

		if strings.Contains(file.Name(), "~$"){
			continue
		}
		log.Printf("start build %s\n", file.Name())
		if excel2Json(file.Name()) < 0{
			log.Printf("%s build fail\n", file.Name())
		}else{
			log.Printf("%s build success\n", file.Name())
		}


	}
	return 0
}

func main()  {
	res := initLog()
	if res < 0{
		fmt.Println("init Log error")
		return
	}

	res = getConfig()
	if res < 0 {
		log.Printf("Error: get confih error\n")
	}

	run()
}