package main

import (
	"bytes"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func TestDocx2md(t *testing.T) {
	files, err := filepath.Glob("testdata/*.docx")
	if err != nil {
		t.Fatal(err)
	}
	for _, docxFile := range files {
		t.Run(filepath.Base(docxFile), func(t *testing.T) {
			mdFile := strings.TrimSuffix(docxFile, ".docx") + ".md"
			expected, err := os.ReadFile(mdFile)
			if err != nil {
				t.Fatalf("failed to read expected markdown: %v", err)
			}

			var buf bytes.Buffer
			old := os.Stdout
			r, w, _ := os.Pipe()
			os.Stdout = w

			err = docx2md(docxFile, false)
			if err != nil {
				t.Fatalf("docx2md failed: %v", err)
			}

			w.Close()
			os.Stdout = old
			buf.ReadFrom(r)

			got := buf.String()
			if got != string(expected) {
				t.Errorf("output mismatch for %s", docxFile)
				t.Errorf("got:\n%s", got)
				t.Errorf("want:\n%s", string(expected))
			}
		})
	}
}

func TestEscape(t *testing.T) {
	tests := []struct {
		input  string
		escape string
		want   string
	}{
		{input: `\`, escape: `\`, want: `\\`},
		{input: `\`, escape: ``, want: `\`},
		{input: `\`, escape: `-`, want: `\`},
		{input: `\\`, escape: `\`, want: `\\\\`},
		{input: `\200`, escape: `\`, want: `\\200`},
	}
	for _, test := range tests {
		got := escape(test.input, test.escape)
		if got != test.want {
			t.Fatalf("want %v, but %v:", test.want, got)
		}
	}
}
