package main

import (
	"encoding/csv"
	"fmt"
	"os"
	"strings"
	"time"
)

func main() {
	fmt.Println("Start")

	// 現在のディレクトリ内の全てのcsvファイルを取得
	files, _ := os.ReadDir(".")
	totalFiles := len(files)
	for i, file := range files {
		if strings.HasSuffix(file.Name(), ".csv") {
			fmt.Printf("Processing file %d of %d: %s\n", i+1, totalFiles, file.Name())
			processFile(file.Name())
		}
	}

	fmt.Println("Complete")
}

func processFile(filename string) {
	file, err := os.Open(filename)
	if err != nil {
		fmt.Println("Error:", err)
		return
	}
	defer file.Close()

	reader := csv.NewReader(file)

	// 全レコード読み込み
	records, _ := reader.ReadAll()

	// ヘッダーの保存
	header := records[0]

	// レコードを通話IDと日付に基づいてグループ化
	groupedRecords := make(map[string]map[string][][]string)
	for _, record := range records[1:] {
		callID, date := record[0], record[1]

		if _, ok := groupedRecords[callID]; !ok {
			groupedRecords[callID] = make(map[string][][]string)
		}

		groupedRecords[callID][date] = append(groupedRecords[callID][date], record)
	}

	// 通話IDと日付ごとにファイル書き出し
	for callID, dateRecords := range groupedRecords {
		for date, records := range dateRecords {
			// 日付のパースと整形
			parsedDate, _ := time.Parse("2006-01-02", date)
			year := parsedDate.Format("2006")
			monthDay := parsedDate.Format("20060102")

			// ディレクトリの作成
			dirName := fmt.Sprintf("%s/%s", year, monthDay)
			os.MkdirAll(dirName, os.ModePerm)

			// ファイル名の作成
			outputFilename := fmt.Sprintf("%s/%s_%s.csv", dirName, monthDay, callID)

			// ファイルが存在するかどうか確認
			_, err := os.Stat(outputFilename)
			newFile := os.IsNotExist(err)

			// ファイルの作成 or 追記モードで開く
			outputFile, err := os.OpenFile(outputFilename, os.O_APPEND|os.O_CREATE|os.O_WRONLY, 0644)
			if err != nil {
				fmt.Println("Error:", err)
				continue
			}
			defer outputFile.Close()

			writer := csv.NewWriter(outputFile)

			// 新規ファイルの場合、ヘッダーの書き出し
			if newFile {
				writer.Write(header)
			}

			// レコードの書き出し
			for _, record := range records {
				writer.Write(record)
			}

			writer.Flush()
		}
	}
}
