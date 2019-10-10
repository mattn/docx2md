package main

import (
	"archive/zip"
	"bytes"
	"encoding/base64"
	"encoding/xml"
	"errors"
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

type Relationship struct {
	Text       string `xml:",chardata"`
	ID         string `xml:"Id,attr"`
	Type       string `xml:"Type,attr"`
	Target     string `xml:"Target,attr"`
	TargetMode string `xml:"TargetMode,attr"`
}

type Relationships struct {
	XMLName      xml.Name       `xml:"Relationships"`
	Text         string         `xml:",chardata"`
	Xmlns        string         `xml:"xmlns,attr"`
	Relationship []Relationship `xml:"Relationship"`
}

type Styles struct {
	XMLName   xml.Name `xml:"styles"`
	Text      string   `xml:",chardata"`
	Mc        string   `xml:"mc,attr"`
	R         string   `xml:"r,attr"`
	W         string   `xml:"w,attr"`
	W14       string   `xml:"w14,attr"`
	W15       string   `xml:"w15,attr"`
	W16se     string   `xml:"w16se,attr"`
	Ignorable string   `xml:"Ignorable,attr"`
	Style     []struct {
		Text        string `xml:",chardata"`
		Type        string `xml:"type,attr"`
		Default     string `xml:"default,attr"`
		StyleId     string `xml:"styleId,attr"`
		CustomStyle string `xml:"customStyle,attr"`
		Name        struct {
			Text string `xml:",chardata"`
			Val  string `xml:"val,attr"`
		} `xml:"name"`
		Rsid struct {
			Text string `xml:",chardata"`
			Val  string `xml:"val,attr"`
		} `xml:"rsid"`
		PPr struct {
			Text    string `xml:",chardata"`
			Spacing struct {
				Text     string `xml:",chardata"`
				Line     string `xml:"line,attr"`
				LineRule string `xml:"lineRule,attr"`
			} `xml:"spacing"`
			KeepNext   string `xml:"keepNext"`
			OutlineLvl struct {
				Text string `xml:",chardata"`
				Val  string `xml:"val,attr"`
			} `xml:"outlineLvl"`
			Tabs struct {
				Text string `xml:",chardata"`
				Tab  []struct {
					Text string `xml:",chardata"`
					Val  string `xml:"val,attr"`
					Pos  string `xml:"pos,attr"`
				} `xml:"tab"`
			} `xml:"tabs"`
			SnapToGrid struct {
				Text string `xml:",chardata"`
				Val  string `xml:"val,attr"`
			} `xml:"snapToGrid"`
			Ind struct {
				Text      string `xml:",chardata"`
				LeftChars string `xml:"leftChars,attr"`
				Left      string `xml:"left,attr"`
			} `xml:"ind"`
		} `xml:"pPr"`
		RPr struct {
			Text   string `xml:",chardata"`
			RFonts struct {
				Text          string `xml:",chardata"`
				Ascii         string `xml:"ascii,attr"`
				HAnsi         string `xml:"hAnsi,attr"`
				Cs            string `xml:"cs,attr"`
				AsciiTheme    string `xml:"asciiTheme,attr"`
				EastAsiaTheme string `xml:"eastAsiaTheme,attr"`
				HAnsiTheme    string `xml:"hAnsiTheme,attr"`
				Cstheme       string `xml:"cstheme,attr"`
			} `xml:"rFonts"`
			Kern struct {
				Text string `xml:",chardata"`
				Val  string `xml:"val,attr"`
			} `xml:"kern"`
			Sz struct {
				Text string `xml:",chardata"`
				Val  string `xml:"val,attr"`
			} `xml:"sz"`
			Lang struct {
				Text string `xml:",chardata"`
				Val  string `xml:"val,attr"`
			} `xml:"lang"`
			SzCs struct {
				Text string `xml:",chardata"`
				Val  string `xml:"val,attr"`
			} `xml:"szCs"`
			Color struct {
				Text       string `xml:",chardata"`
				Val        string `xml:"val,attr"`
				ThemeColor string `xml:"themeColor,attr"`
			} `xml:"color"`
			U struct {
				Text string `xml:",chardata"`
				Val  string `xml:"val,attr"`
			} `xml:"u"`
			VertAlign struct {
				Text string `xml:",chardata"`
				Val  string `xml:"val,attr"`
			} `xml:"vertAlign"`
		} `xml:"rPr"`
		BasedOn struct {
			Text string `xml:",chardata"`
			Val  string `xml:"val,attr"`
		} `xml:"basedOn"`
		Next struct {
			Text string `xml:",chardata"`
			Val  string `xml:"val,attr"`
		} `xml:"next"`
		Link struct {
			Text string `xml:",chardata"`
			Val  string `xml:"val,attr"`
		} `xml:"link"`
		UiPriority struct {
			Text string `xml:",chardata"`
			Val  string `xml:"val,attr"`
		} `xml:"uiPriority"`
		QFormat        string `xml:"qFormat"`
		UnhideWhenUsed string `xml:"unhideWhenUsed"`
		SemiHidden     string `xml:"semiHidden"`
		TblPr          struct {
			Text   string `xml:",chardata"`
			TblInd struct {
				Text string `xml:",chardata"`
				W    string `xml:"w,attr"`
				Type string `xml:"type,attr"`
			} `xml:"tblInd"`
			TblCellMar struct {
				Text string `xml:",chardata"`
				Top  struct {
					Text string `xml:",chardata"`
					W    string `xml:"w,attr"`
					Type string `xml:"type,attr"`
				} `xml:"top"`
				Left struct {
					Text string `xml:",chardata"`
					W    string `xml:"w,attr"`
					Type string `xml:"type,attr"`
				} `xml:"left"`
				Bottom struct {
					Text string `xml:",chardata"`
					W    string `xml:"w,attr"`
					Type string `xml:"type,attr"`
				} `xml:"bottom"`
				Right struct {
					Text string `xml:",chardata"`
					W    string `xml:"w,attr"`
					Type string `xml:"type,attr"`
				} `xml:"right"`
			} `xml:"tblCellMar"`
			TblBorders struct {
				Text string `xml:",chardata"`
				Top  struct {
					Text  string `xml:",chardata"`
					Val   string `xml:"val,attr"`
					Sz    string `xml:"sz,attr"`
					Space string `xml:"space,attr"`
					Color string `xml:"color,attr"`
				} `xml:"top"`
				Left struct {
					Text  string `xml:",chardata"`
					Val   string `xml:"val,attr"`
					Sz    string `xml:"sz,attr"`
					Space string `xml:"space,attr"`
					Color string `xml:"color,attr"`
				} `xml:"left"`
				Bottom struct {
					Text  string `xml:",chardata"`
					Val   string `xml:"val,attr"`
					Sz    string `xml:"sz,attr"`
					Space string `xml:"space,attr"`
					Color string `xml:"color,attr"`
				} `xml:"bottom"`
				Right struct {
					Text  string `xml:",chardata"`
					Val   string `xml:"val,attr"`
					Sz    string `xml:"sz,attr"`
					Space string `xml:"space,attr"`
					Color string `xml:"color,attr"`
				} `xml:"right"`
				InsideH struct {
					Text  string `xml:",chardata"`
					Val   string `xml:"val,attr"`
					Sz    string `xml:"sz,attr"`
					Space string `xml:"space,attr"`
					Color string `xml:"color,attr"`
				} `xml:"insideH"`
				InsideV struct {
					Text  string `xml:",chardata"`
					Val   string `xml:"val,attr"`
					Sz    string `xml:"sz,attr"`
					Space string `xml:"space,attr"`
					Color string `xml:"color,attr"`
				} `xml:"insideV"`
			} `xml:"tblBorders"`
		} `xml:"tblPr"`
	} `xml:"style"`
}

type file struct {
	rels   Relationships
	styles Styles
	r      *zip.ReadCloser
	embed  bool
	num    int
}

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

func escape(s, set string) string {
	replacer := []string{}
	for _, r := range []rune(set) {
		rs := string(r)
		replacer = append(replacer, rs, `\`+rs)
	}
	return strings.NewReplacer(replacer...).Replace(s)
}

func (zf *file) extract(rel *Relationship, w io.Writer) error {
	err := os.MkdirAll(filepath.Dir(rel.Target), 0755)
	if err != nil {
		return err
	}
	for _, f := range zf.r.File {
		if f.Name != "word/"+rel.Target {
			continue
		}
		rc, err := f.Open()
		if err != nil {
			return err
		}
		defer rc.Close()

		b := make([]byte, f.UncompressedSize64)
		n, err := rc.Read(b)
		if err != nil && err != io.EOF {
			return err
		}
		if zf.embed {
			fmt.Fprintf(w, "![](data:image/png;base64,%s)",
				base64.StdEncoding.EncodeToString(b[:n]))
		} else {
			err = ioutil.WriteFile(rel.Target, b, 0644)
			if err != nil {
				return err
			}
			fmt.Fprintf(w, "![](%s)", escape(rel.Target, "()"))
		}
		break
	}
	return nil
}

func attr(attrs []xml.Attr, name string) (string, bool) {
	for _, attr := range attrs {
		if attr.Name.Local == name {
			return attr.Value, true
		}
	}
	return "", false
}

func (zf *file) walk(node *Node, w io.Writer) error {
	switch node.XMLName.Local {
	case "hyperlink":
		fmt.Fprint(w, "[")
		var cbuf bytes.Buffer
		for _, n := range node.Nodes {
			if err := zf.walk(&n, &cbuf); err != nil {
				return err
			}
		}
		fmt.Fprint(w, escape(cbuf.String(), "[]"))
		fmt.Fprint(w, "]")

		fmt.Fprint(w, "(")
		if id, ok := attr(node.Attrs, "id"); ok {
			for _, rel := range zf.rels.Relationship {
				if id == rel.ID {
					fmt.Fprint(w, escape(rel.Target, "()"))
					break
				}
			}
		}
		fmt.Fprint(w, ")")
	case "t":
		fmt.Fprint(w, string(node.Content))
	case "pPr":
		code := false
		for _, n := range node.Nodes {
			switch n.XMLName.Local {
			case "ind":
				if left, ok := attr(n.Attrs, "left"); ok {
					if i, err := strconv.Atoi(left); err == nil {
						fmt.Fprint(w, strings.Repeat(" ", i/360))
					}
				}
			case "pStyle":
				if val, ok := attr(n.Attrs, "val"); ok {
					if strings.HasPrefix(val, "Heading") {
						if i, err := strconv.Atoi(val[7:]); err == nil {
							fmt.Fprint(w, strings.Repeat("#", i)+" ")
						}
					} else if val == "Code" {
						code = true
					} else {
						for _, style := range zf.styles.Style {
							if val == style.StyleId {
							}
						}
						if i, err := strconv.Atoi(val); err == nil {
							fmt.Fprint(w, strings.Repeat("#", i)+" ")
						}
					}
				}
			}
		}
		if code {
			fmt.Fprint(w, "`")
		}
		for _, n := range node.Nodes {
			if err := zf.walk(&n, w); err != nil {
				return err
			}
		}
		if code {
			fmt.Fprint(w, "`")
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
				if err := zf.walk(&tc, &cbuf); err != nil {
					return err
				}
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
					fmt.Fprint(w, escape(row[j], "|"))
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
			if n.XMLName.Local != "rPr" {
				continue
			}
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
		if strike {
			fmt.Fprint(w, "~~")
		}
		if bold {
			fmt.Fprint(w, "**")
		}
		if italic {
			fmt.Fprint(w, "*")
		}
		var cbuf bytes.Buffer
		for _, n := range node.Nodes {
			if err := zf.walk(&n, &cbuf); err != nil {
				return err
			}
		}
		fmt.Fprint(w, escape(cbuf.String(), `*~\`))
		if italic {
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
			if err := zf.walk(&n, w); err != nil {
				return err
			}
		}
		fmt.Fprintln(w)
	case "blip":
		if id, ok := attr(node.Attrs, "embed"); ok {
			for _, rel := range zf.rels.Relationship {
				if id != rel.ID {
					continue
				}
				if err := zf.extract(&rel, w); err != nil {
					return err
				}
			}
		}
	case "Fallback":
	case "txbxContent":
		var cbuf bytes.Buffer
		for _, n := range node.Nodes {
			if err := zf.walk(&n, &cbuf); err != nil {
				return err
			}
		}
		fmt.Fprintln(w, "\n```\n"+cbuf.String()+"```")
	default:
		for _, n := range node.Nodes {
			if err := zf.walk(&n, w); err != nil {
				return err
			}
		}
	}

	return nil
}

func readFile(f *zip.File) (*Node, error) {
	rc, err := f.Open()
	defer rc.Close()

	b, _ := ioutil.ReadAll(rc)
	if err != nil {
		return nil, err
	}

	var node Node
	err = xml.Unmarshal(b, &node)
	if err != nil {
		return nil, err
	}
	return &node, nil
}

func findFile(files []*zip.File, target string) *zip.File {
	for _, f := range files {
		if f.Name == target {
			return f
		}
	}
	return nil
}

func docx2md(arg string, embed bool) error {
	r, err := zip.OpenReader(arg)
	if err != nil {
		return err
	}
	defer r.Close()

	var rels Relationships
	var styles Styles

	for _, f := range r.File {
		switch f.Name {
		case "word/_rels/document.xml.rels":
			rc, err := f.Open()
			defer rc.Close()

			b, _ := ioutil.ReadAll(rc)
			if err != nil {
				return err
			}

			err = xml.Unmarshal(b, &rels)
			if err != nil {
				return err
			}
		case "word/styles.xml":
			rc, err := f.Open()
			defer rc.Close()

			b, _ := ioutil.ReadAll(rc)
			if err != nil {
				return err
			}

			err = xml.Unmarshal(b, &styles)
			if err != nil {
				return err
			}
		}
	}

	f := findFile(r.File, "word/document.xml")
	if f == nil {
		return errors.New("incorrect document")
	}
	node, err := readFile(f)
	if err != nil {
		return err
	}

	var buf bytes.Buffer
	zf := &file{
		r:      r,
		rels:   rels,
		styles: styles,
		embed:  embed,
		num:    0,
	}
	err = zf.walk(node, &buf)
	if err != nil {
		return err
	}
	fmt.Print(buf.String())

	return nil
}

func main() {
	var embed bool
	flag.BoolVar(&embed, "embed", false, "embed resources")
	flag.Parse()
	for _, arg := range flag.Args() {
		if err := docx2md(arg, embed); err != nil {
			log.Fatal(err)
		}
	}
}
