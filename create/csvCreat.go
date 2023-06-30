package create

import (
	"encoding/csv"
	"fmt"
	"math/rand"
	"os"
	"time"
)

const (
	numRows   = 1000
	startYear = 2020
	endYear   = 2023
)

func main() {
	rand.Seed(time.Now().UnixNano())

	// テストデータ生成
	data := make([][]string, numRows+1)

	// ヘッダ行
	data[0] = []string{"通話ID", "日付", "内容"}

	// データ行
	for i := 1; i <= numRows; i++ {
		// ランダムな通話ID（AまたはB）
		callID := "A"
		if rand.Intn(2) == 1 {
			callID = "B"
		}

		// ランダムな日付
		date := randomDate(startYear, endYear)

		// ランダムな内容
		content := fmt.Sprintf("test%d", i)

		data[i] = []string{callID, date, content}
	}

	// CSVファイル作成
	file, err := os.Create("test.csv")
	if err != nil {
		fmt.Println("Error:", err)
		return
	}
	defer file.Close()

	writer := csv.NewWriter(file)
	writer.WriteAll(data) // calls Flush internally
}

// randomDate returns a random date string between startYear and endYear.
func randomDate(startYear, endYear int) string {
	start := time.Date(startYear, 1, 1, 0, 0, 0, 0, time.UTC)
	end := time.Date(endYear, 12, 31, 23, 59, 59, 0, time.UTC)

	delta := int(end.Sub(start).Hours() / 24)
	incr := rand.Intn(delta)

	randDate := start.AddDate(0, 0, incr)
	return randDate.Format("2006-01-02")
}
