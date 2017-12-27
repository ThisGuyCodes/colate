package main

import (
	"bytes"
	"encoding/csv"
	"flag"
	"io/ioutil"
	"log"
	"path/filepath"
	"strconv"
	"strings"
	"text/template"

	"github.com/360EntSecGroup-Skylar/excelize"
	"go.uber.org/zap"
	"go.uber.org/zap/zapcore"
)

const ()

var (
	fileDir    = flag.String("dir", ".", "the directory of *.xlsx files you want to process")
	outputName = flag.String("output", "./output.xlsx", "file to output the results to")
	sheetName  = flag.String("sheet", "", "sheet name to pull data from")
	rowStart   = flag.Int("rowStart", 0, "row data starts on, to account for headers")
	rowCount   = flag.Int("rowCount", 0, "number of rows to take, 0 for no limit")

	newColumns = flag.String("columns", "{{.FileName}}", "new columns to prepend, comma-seperated with template syntax")

	logV = flag.Bool("v", false, "enable verbose logging. Like no really, super verbose")

	logConfig = zap.Config{
		Level:             zap.NewAtomicLevelAt(zap.DebugLevel),
		Development:       true,
		DisableCaller:     true,
		DisableStacktrace: true,
		Sampling:          nil,
		Encoding:          "json",
		EncoderConfig: zapcore.EncoderConfig{
			TimeKey:        "ts",
			LevelKey:       "level",
			NameKey:        "logger",
			CallerKey:      "caller",
			MessageKey:     "msg",
			StacktraceKey:  "stacktrace",
			LineEnding:     zapcore.DefaultLineEnding,
			EncodeLevel:    zapcore.CapitalLevelEncoder,
			EncodeTime:     zapcore.EpochTimeEncoder,
			EncodeDuration: zapcore.SecondsDurationEncoder,
			EncodeCaller:   zapcore.ShortCallerEncoder,
		},
		OutputPaths:      []string{"stdout"},
		ErrorOutputPaths: []string{"stderr"},
	}
)

func getColumns(columnString string) ([]*template.Template, error) {
	r := csv.NewReader(bytes.NewBufferString(columnString))
	columns, err := r.Read()
	if err != nil {
		return []*template.Template{}, err
	}
	templates := make([]*template.Template, len(columns))
	for i := range templates {
		templates[i], err = template.New(strconv.Itoa(i)).Parse(columns[i])
		if err != nil {
			return []*template.Template{}, err
		}
	}
	return templates, nil
}

type ColumnData struct {
	Cells    []string
	FileName string
	RowNum   int
}

func main() {
	flag.Parse()

	var err error
	var l *zap.SugaredLogger
	if *logV {
		var ll *zap.Logger
		ll, err = logConfig.Build()
		l = ll.Sugar()
	} else {
		var ll *zap.Logger
		logConfig.Level = zap.NewAtomicLevelAt(zap.WarnLevel)
		ll, err = logConfig.Build()
		l = ll.Sugar()
	}
	if err != nil {
		log.Fatalf("can't initialize zap logger: %v", err)
	}
	defer l.Sync()
	l = l.Named("colate")

	columns, err := getColumns(*newColumns)
	if err != nil {
		l.Fatalw("couldn't parse columns flag",
			"newColumns", *newColumns,
		)
	}

	files, err := listFiles(*fileDir)

	l.Debugw("files",
		"files", files,
	)

	if err != nil {
		l.Fatalw("fatal error",
			"error", err,
		)
	}

	var data [][]string
	for _, file := range files {
		l := l.With(
			"file", file,
		)
		l.Debug("starting file")

		fileBase := filepath.Base(file)
		thisData, err := getRows(l, file, *sheetName, *rowStart, *rowCount)
		if err != nil {
			l.Fatalw("fatal error",
				"error", err,
			)
		}

		// prepend the file name
		columnData := ColumnData{
			FileName: fileBase,
		}
		thisData = prependColumns(l, thisData, columnData, columns)

		// add to output data
		data = append(data, thisData...)
	}

	// save data to new file
	output := createFile(l, *sheetName, data)
	// write out!
	err = output.SaveAs(*outputName)
	if err != nil {
		l.Fatalw("fatal error",
			"error", err,
		)
	}
}

func listFiles(dir string) ([]string, error) {
	files, err := ioutil.ReadDir(dir)
	if err != nil {
		return []string{}, err
	}

	var ret []string
	for _, file := range files {
		if ok, _ := filepath.Match("*.xlsx", strings.ToLower(file.Name())); ok {
			ret = append(ret, filepath.Join(dir, file.Name()))
		}
	}
	return ret, nil
}

func createFile(l *zap.SugaredLogger, name string, data [][]string) *excelize.File {
	f := excelize.NewFile()
	sheetIndex := f.NewSheet(name)
	f.SetActiveSheet(sheetIndex)
	writeData(l, f, 0, name, data)
	return f
}

func writeData(l *zap.SugaredLogger, f *excelize.File, startRow int, sheet string, data [][]string) int {
	l.Debugw("writeData()",
		"f", "<omitted>",
		"startRow", startRow,
		"sheet", sheet,
		"data", len(data),
	)
	for ri, row := range data {
		for ci, value := range row {
			// construct cell name. Note: excel is 1 indexed
			loc := excelize.ToAlphaString(ci) + strconv.Itoa(startRow+ri+1)
			f.SetCellStr(sheet, loc, value)
		}
	}

	return startRow + len(data)
}

// prependColumns will take a two dimensional slice of strings and insert a new
// column. The values for the new column are taken from "source", the input
// slice is repeated as many times as necessary to fill all rows of the input
func prependColumns(l *zap.SugaredLogger, rows [][]string, columnData ColumnData, sources []*template.Template) [][]string {
	l.Debugw("insertColumn()",
		"rows", len(rows),
		"columnData", columnData,
	)
	for ri, row := range rows {
		columnData.Cells = row
		columnData.RowNum = ri
		newData := make([]string, len(sources))
		for si, source := range sources {
			toPut := new(bytes.Buffer)
			// get the value to put
			err := source.Execute(toPut, columnData)
			if err != nil {
				l.Warnw("error executing column template",
					"err", err,
					"template", source.DefinedTemplates(),
					"columnData", columnData,
				)
			}
			newData[si] = toPut.String()
		}

		// reconstruct the new slice, replacing the current row
		rows[ri] = append(newData, row...)
	}
	return rows
}

func getRows(l *zap.SugaredLogger, file, sheet string, start, count int) ([][]string, error) {
	f, err := excelize.OpenFile(file)
	if err != nil {
		return [][]string{}, err
	}

	rows := f.GetRows(sheet)

	// start from intended position
	rows = rows[start:]

	// fill in blanks with preceeding values
	for ri, row := range rows {
		for ci, cell := range row {
			if cell == "" && ri != 0 {
				l.Debugw("inheriting empty cell value from previous row",
					"row", ri,
					"column", ci,
					"value", cell,
					"inherited", rows[ri-1][ci],
				)
				row[ci] = rows[ri-1][ci]
			}
		}
		count = count - 1
		if count == 0 {
			break
		}
	}

	return rows, nil
}
