package main

import (
	"encoding/json"
	"fmt"
	"github.com/xuri/excelize/v2"
	"io"
	"log"
	"net/http"
	"strings"
)

type ExportRequest struct {
	Date    string   `json:"date"`
	Records []Record `json:"records"`
}

type Record struct {
	WorkKbn     string `json:"lbWorkKbn"`
	StartTime   string `json:"lbStartTime"`
	EndTime     string `json:"lbEndTime"`
	OtStartTime string `json:"lbOtStartTime"`
	OtEndTime   string `json:"lbOtEndTime"`
	Note        string `json:"lbNote"`
}

func main() {

	http.HandleFunc("/export", func(w http.ResponseWriter, r *http.Request) {
		// HTTPメソッドをチェック（POSTのみ許可）
		if r.Method != http.MethodPost {
			w.WriteHeader(http.StatusMethodNotAllowed) // 405
			w.Write([]byte("POSTだけだよー"))
			return
		}

		body, _ := io.ReadAll(r.Body)
		var exportRequest ExportRequest
		if err := json.Unmarshal(body, &exportRequest); err != nil {
			fmt.Println(err)
			return
		}

		f, err := excelize.OpenFile("template.xlsx")
		if err != nil {
			fmt.Println(err)
			return
		}
		defer func() {
			if err := f.Close(); err != nil {
				fmt.Println(err)
			}
		}()

		f.SetCellValue("勤務表-2018MM", "A3", exportRequest.Date)
		for index, record := range exportRequest.Records {
			f.SetCellValue("勤務表-2018MM", fmt.Sprintf("C%d", 16+index), record.WorkKbn)
			f.SetCellValue("勤務表-2018MM", fmt.Sprintf("D%d", 16+index), record.StartTime)
			f.SetCellValue("勤務表-2018MM", fmt.Sprintf("E%d", 16+index), record.EndTime)
			f.SetCellValue("勤務表-2018MM", fmt.Sprintf("G%d", 16+index), record.OtStartTime)
			f.SetCellValue("勤務表-2018MM", fmt.Sprintf("H%d", 16+index), record.OtEndTime)
			f.SetCellValue("勤務表-2018MM", fmt.Sprintf("K%d", 16+index), record.Note)
		}

		splitDate := strings.Split(exportRequest.Date, "/")
		f.SetSheetName("勤務表-2018MM", fmt.Sprintf("勤務表-%s%s", splitDate[0], splitDate[1]))
		w.Header().Set("Content-Type", "application/octet-stream")
		w.Header().Set("Content-Disposition", fmt.Sprintf("attachment; filename=%s年%s月分の作業報告書(DUY).xlsx", splitDate[0], splitDate[1]))
		//files, err := ioutil.ReadFile(result)
		//w.Write(files)
		buf, _ := f.WriteToBuffer()

		w.Write(buf.Bytes())
	})
	log.Fatal(http.ListenAndServe(":8080", nil))
}
