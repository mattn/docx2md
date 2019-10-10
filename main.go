package main

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"flag"
	"fmt"
	"io"
	"io/ioutil"
	"log"
	"os"
	"path/filepath"
	"strconv"
	"strings"

	"github.com/mattn/go-runewidth"
)

type Relationships struct {
	XMLName      xml.Name `xml:"Relationships"`
	Text         string   `xml:",chardata"`
	Xmlns        string   `xml:"xmlns,attr"`
	Relationship []struct {
		Text            string `xml:",chardata"`
		ID              string `xml:"Id,attr"`
		Type            string `xml:"Type,attr"`
		Target          string `xml:"Target,attr"`
		TargetMode      string `xml:"TargetMode,attr"`
		mustBeExtracted bool
	} `xml:"Relationship"`
}

var rels Relationships

type Node struct {
	XMLName xml.Name
	Attrs   []xml.Attr `xml:"-"`
	Content []byte     `xml:",innerxml"`
	Nodes   []Node     `xml:",any"`
}

func (n *Node) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	n.Attrs = start.Attr
	type node Node

	return d.DecodeElement((*node)(n), &start)
}

func walk(node Node, w io.Writer) {
	switch node.XMLName.Local {
	case "hyperlink":
		fmt.Fprint(w, "[")
		for _, n := range node.Nodes {
			walk(n, w)
		}
		fmt.Fprint(w, "]")

		fmt.Fprint(w, "(")
		for _, attr := range node.Attrs {
			if attr.Name.Local == "id" {
				for _, rel := range rels.Relationship {
					if attr.Value == rel.ID {
						fmt.Fprint(w, rel.Target)
					}
				}
			}
		}
		fmt.Fprint(w, ")")
	case "t":
		fmt.Fprint(w, string(node.Content))
	case "pPr":
		for _, n := range node.Nodes {
			switch n.XMLName.Local {
			case "ind":
				for _, attr := range n.Attrs {
					if attr.Name.Local == "left" {
						if i, err := strconv.Atoi(attr.Value); err == nil {
							fmt.Fprint(w, strings.Repeat(" ", i/360))
						}
					}
				}
			case "pStyle":
				for _, attr := range n.Attrs {
					if attr.Name.Local == "val" {
						if i, err := strconv.Atoi(attr.Value); err == nil {
							fmt.Fprint(w, strings.Repeat("#", i)+" ")
						}
					}
				}
			}
		}
		for _, n := range node.Nodes {
			walk(n, w)
		}
	case "tbl":
		var rows [][]string
		for _, tr := range node.Nodes {
			if tr.XMLName.Local != "tr" {
				continue
			}
			var cols []string
			for _, tc := range tr.Nodes {
				if tc.XMLName.Local != "tc" {
					continue
				}
				var cbuf bytes.Buffer
				walk(tc, &cbuf)
				cols = append(cols, strings.Replace(cbuf.String(), "\n", "", -1))
			}
			rows = append(rows, cols)
		}
		maxcol := 0
		for _, cols := range rows {
			if len(cols) > maxcol {
				maxcol = len(cols)
			}
		}
		widths := make([]int, maxcol)
		for _, row := range rows {
			for i := 0; i < maxcol; i++ {
				if i < len(row) {
					width := runewidth.StringWidth(row[i])
					if widths[i] < width {
						widths[i] = width
					}
				}
			}
		}
		for i, row := range rows {
			if i == 0 {
				for j := 0; j < maxcol; j++ {
					fmt.Fprint(w, "|")
					fmt.Fprint(w, strings.Repeat(" ", widths[j]))
				}
				fmt.Fprint(w, "|\n")
				for j := 0; j < maxcol; j++ {
					fmt.Fprint(w, "|")
					fmt.Fprint(w, strings.Repeat("-", widths[j]))
				}
				fmt.Fprint(w, "|\n")
			}
			for j := 0; j < maxcol; j++ {
				fmt.Fprint(w, "|")
				if j < len(row) {
					width := runewidth.StringWidth(row[j])
					fmt.Fprint(w, row[j])
					fmt.Fprint(w, strings.Repeat(" ", widths[j]-width))
				} else {
					fmt.Fprint(w, strings.Repeat(" ", widths[j]))
				}
			}
			fmt.Fprint(w, "|\n")
		}
		fmt.Fprint(w, "\n")
	case "numPr":
		fmt.Fprint(w, "* ")
	case "r":
		bold := false
		italic := false
		strike := false
		for _, n := range node.Nodes {
			if n.XMLName.Local == "rPr" {
				for _, nn := range n.Nodes {
					switch nn.XMLName.Local {
					case "b":
						bold = true
					case "i":
						italic = true
					case "strike":
						strike = true
					}
				}
			}
		}
		if bold {
			fmt.Fprint(w, "**")
		}
		if italic {
			fmt.Fprint(w, "*")
		}
		if strike {
			fmt.Fprint(w, "~~")
		}
		for _, n := range node.Nodes {
			walk(n, w)
		}
		if bold {
			fmt.Fprint(w, "*")
		}
		if bold {
			fmt.Fprint(w, "**")
		}
		if strike {
			fmt.Fprint(w, "~~")
		}
	case "p":
		for _, n := range node.Nodes {
			walk(n, w)
		}
		fmt.Fprintln(w)
	case "blip":
		for _, attr := range node.Attrs {
			if attr.Name.Local == "embed" {
				for i, rel := range rels.Relationship {
					if attr.Value == rel.ID {
						fmt.Fprintf(w, "![](%s)", rel.Target)
						rels.Relationship[i].mustBeExtracted = true
					}
				}
			}
		}
	case "Fallback":
	case "txbxContent":
		var cbuf bytes.Buffer
		for _, n := range node.Nodes {
			walk(n, &cbuf)
		}
		fmt.Fprintln(w, "\n```\n"+cbuf.String()+"```")
	default:
		for _, n := range node.Nodes {
			walk(n, w)
		}
	}
}

func docx2txt(arg string) {
	r, err := zip.OpenReader(arg)
	if err != nil {
		log.Fatal(err)
	}
	defer r.Close()

	for _, f := range r.File {
		if f.Name == "word/_rels/document.xml.rels" {
			rc, err := f.Open()
			defer rc.Close()

			b, _ := ioutil.ReadAll(rc)
			if err != nil {
				log.Fatal(err)
			}

			err = xml.Unmarshal(b, &rels)
			if err != nil {
				log.Fatal(err)
			}
		}
	}

	for _, f := range r.File {
		if f.Name == "word/document.xml" {
			rc, err := f.Open()
			defer rc.Close() // TODO do not call defer in loop

			b, _ := ioutil.ReadAll(rc)
			if err != nil {
				log.Fatal(err)
			}

			var node Node
			err = xml.Unmarshal(b, &node)
			if err != nil {
				log.Fatal(err)
			}
			var buf bytes.Buffer
			walk(node, &buf)
			fmt.Print(buf.String())
		}
	}

	for _, rel := range rels.Relationship {
		if rel.mustBeExtracted {
			err = os.MkdirAll(filepath.Dir(rel.Target), 0755)
			if err != nil {
				log.Fatal(err)
			}
			for _, f := range r.File {
				if f.Name == "word/"+rel.Target {
					rc, err := f.Open()
					if err != nil {
						log.Fatal(err)
					}
					defer rc.Close() // TODO do not call defer in loop

					b := make([]byte, f.UncompressedSize64)
					_, err = rc.Read(b)
					if err != nil {
						log.Fatal(err)
					}
					err = ioutil.WriteFile(rel.Target, b, 0644)
					if err != nil {
						log.Fatal(err)
					}
				}
			}
		}
	}
}

func main() {
	flag.Parse()
	for _, arg := range flag.Args() {
		docx2txt(arg)
	}
}
