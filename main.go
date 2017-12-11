package main

import (
	"flag"
	"io/ioutil"
	"log"
	"path/filepath"
	"strconv"
	"strings"

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
		thisData, err := getRows(l, file, *sheetName, *rowStart)
		if err != nil {
			l.Fatalw("fatal error",
				"error", err,
			)
		}

		// prepend the file name
		thisData = insertColumn(thisData, 0, []string{fileBase})

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
	for ri, row := range data {
		for ci, value := range row {
			// construct the cell name. Note: excel is 1 indexed
			loc := excelize.ToAlphaString(ci) + strconv.Itoa(ri+1)
			l.Debugw("output cell value",
				"sheetName", name,
				"cellLocation", loc,
				"value", value,
			)
			f.SetCellStr(name, loc, value)
		}
	}
	return f
}

// insertColumn will take a two dimensional slice of strings and insert a new
// column. The values for the new column are taken from "source", the input
// slice is repeated as many times as necessary to fill all rows of the input
func insertColumn(rows [][]string, position int, source []string) [][]string {
	sourceLen := len(source)
	for ri, row := range rows {
		// get the value to put
		toPut := source[ri%sourceLen]
		// reconstruct the new slice, replacing the current row
		rows[ri] = append(row[:position], append([]string{toPut}, row[position:]...)...)
	}
	return rows
}

func getRows(l *zap.SugaredLogger, file, sheet string, start int) ([][]string, error) {
	f, err := excelize.OpenFile(file)
	if err != nil {
		return [][]string{}, err
	}

	rows := f.GetRows(sheet)

	// start from intended position
	rows = rows[start:]

	// fill in blanks with preceeding values
	for ri, row := range rows {
		l.Debugw("Starting next row",
			"row", ri,
			"length", len(row),
		)
		for ci, cell := range row {
			l.Debugw("cell value",
				"row", ri,
				"column", ci,
				"value", cell,
			)
			if cell == "" && ri != 0 {
				l.Infow("inheriting empty cell value from previous row",
					"row", ri,
					"column", ci,
					"value", cell,
					"inherited", rows[ri-1][ci],
				)
				row[ci] = rows[ri-1][ci]
			}
		}
	}

	return rows, nil
}
