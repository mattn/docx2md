package main

import (
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
