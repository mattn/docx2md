package docx2md

import (
	"fmt"
	"io/ioutil"
	"testing"
)

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

func TestConversion(t *testing.T) {

	v, err := ioutil.ReadFile("./tests.docx") //read the content of file
	if err != nil {
		fmt.Println(err)
		return
	}

	mk, err := Docx2md(v, false)
	if err != nil {
		return
	}

	fmt.Println(mk)

}
