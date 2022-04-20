package main

import (
	"fmt"
	"regexp"
	"strconv"
	"strings"

	"github.com/gingfrederik/docx"
	"github.com/xuri/excelize/v2"
)

func CreatWord(str string, title string) {
	w := docx.NewFile()
	for _, text := range strings.Split(str, "\n") {
		para := w.AddParagraph()
		para.AddText(text).Size(8)
	}
	w.Save(title)
}
func MatchOrNot(str1 string, str2 string) bool {
	match, err := regexp.MatchString(str1, str2)
	if err != nil {
		return false
	} else {
		return match
	}
}

func GetNeedColumn(arr []string) []int {
	var (
		question []int
		choice   []int
		answer   []int
		content  []int
	)
	for index, item := range arr {
		if MatchOrNot("题目", item) {
			question = append(question, index)
		}
		if MatchOrNot("(?i)A|B|C|D|E|F", item) {
			choice = append(choice, index)
		}
		if MatchOrNot("答案", item) {
			answer = append(answer, index)
		}
	}
	if len(question) > 0 {
		content = append(content, question[0])
	}
	if len(choice) > 0 {
		content = append(content, choice...)
	}
	if len(answer) > 0 {
		content = append(content, answer[0])
	}
	return content
}

func JoinList(arr []string) string {
	letterlist := []string{"A", "B", "C", "D", "E", "F"}
	arr[0] = strings.TrimSpace(arr[0]) + "\n"
	if len(arr) == 2 {
		arr[len(arr)-1] = " 答案：" + strings.TrimSpace(arr[len(arr)-1]) + "\n"
	}
	if len(arr) > 2 {
		for index := range arr {
			if (index >= 1) && (index < len(arr)-1) && (arr[index] != "") && (strings.TrimSpace(arr[index]) != "") {
				arr[index] = letterlist[index-1] + "." + strings.TrimSpace(arr[index]) + " "
			}
		}
		arr[len(arr)-1] = " 答案：" + strings.TrimSpace(arr[len(arr)-1]) + "\n"
	}
	return strings.Join(arr, "")
}

func main() {
	f, err := excelize.OpenFile("safety.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	for index, item := range f.GetSheetList() {
		rows, err := f.GetRows(item)
		if err != nil {
			fmt.Println(err)
			return
		}
		sum := []string{}
		data := []int{}
		for indexinner, row := range rows {
			if indexinner == 0 {
				data = GetNeedColumn(row)
			}
			if len(data) == 0 {
				break
			}
			if indexinner > 0 && len(data) > 0 {
				everyrow := []string{}
				var questionindex string = "(" + strconv.Itoa(indexinner) + ")"
				for _, num := range data {
					everyrow = append(everyrow, row[num])
				}
				var everyrowlist string = questionindex + JoinList(everyrow)
				sum = append(sum, everyrowlist)
			}
		}
		if len(sum) > 0 {
			var total string = strings.Join(sum, "")
			var title string = f.GetSheetName(index) + "go.docx"
			CreatWord(total, title)
		}
	}
}
