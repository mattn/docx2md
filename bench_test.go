package main

import (
	"os"
	"testing"
)

func BenchmarkDocx2md(b *testing.B) {
	path := os.Getenv("DOCX2MD_BENCH_FILE")
	if path == "" {
		b.Skip("DOCX2MD_BENCH_FILE not set")
	}
	null, err := os.OpenFile(os.DevNull, os.O_WRONLY, 0)
	if err != nil {
		b.Fatal(err)
	}
	defer null.Close()
	old := os.Stdout
	os.Stdout = null
	defer func() { os.Stdout = old }()
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		if err := docx2md(path, Config{}); err != nil {
			b.Fatal(err)
		}
	}
}
