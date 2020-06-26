package csv2xlsx

import (
	"encoding/csv"
	"io"

	xlsx "github.com/plandem/xlsx"
)

// Convert will convert a CSV to an xlsx.Spreadsheet.
func Convert(r io.Reader, sheetName string) (*xlsx.Spreadsheet, error) {
	reader := csv.NewReader(r)
	reader.Comma = ','
	reader.Comment = '#'
	reader.LazyQuotes = true
	rows, err := reader.ReadAll()
	if err != nil {
		return nil, err
	}

	xl := xlsx.New()
	xlSheet := xl.AddSheet(sheetName)
	defer xlSheet.Close()

	for iRow, rowVals := range rows {
		xlSheet.InsertRow(iRow)
		for iColumn, val := range rowVals {
			xlSheet.Cell(iColumn, iRow).SetValue(val)
		}
	}

	return xl, nil
}
