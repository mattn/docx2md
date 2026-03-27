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
	"log"
	"os"
	"path"
	"path/filepath"
	"runtime"
	"strconv"
	"strings"

	"github.com/mattn/go-runewidth"
)

const name = "docx2md"
const version = "0.0.13"

var revision = "HEAD"

// Relationship is
type Relationship struct {
	Text       string `xml:",chardata"`
	ID         string `xml:"Id,attr"`
	Type       string `xml:"Type,attr"`
	Target     string `xml:"Target,attr"`
	TargetMode string `xml:"TargetMode,attr"`
}

// Relationships is
type Relationships struct {
	XMLName      xml.Name       `xml:"Relationships"`
	Text         string         `xml:",chardata"`
	Xmlns        string         `xml:"xmlns,attr"`
	Relationship []Relationship `xml:"Relationship"`
}

// TextVal is
type TextVal struct {
	Text string `xml:",chardata"`
	Val  string `xml:"val,attr"`
}

// NumberingLvl is
type NumberingLvl struct {
	Text      string  `xml:",chardata"`
	Ilvl      string  `xml:"ilvl,attr"`
	Tplc      string  `xml:"tplc,attr"`
	Tentative string  `xml:"tentative,attr"`
	Start     TextVal `xml:"start"`
	NumFmt    TextVal `xml:"numFmt"`
	LvlText   TextVal `xml:"lvlText"`
	LvlJc     TextVal `xml:"lvlJc"`
	PPr       struct {
		Text string `xml:",chardata"`
		Ind  struct {
			Text    string `xml:",chardata"`
			Left    string `xml:"left,attr"`
			Hanging string `xml:"hanging,attr"`
		} `xml:"ind"`
	} `xml:"pPr"`
	RPr struct {
		Text string `xml:",chardata"`
		U    struct {
			Text string `xml:",chardata"`
			Val  string `xml:"val,attr"`
		} `xml:"u"`
		RFonts struct {
			Text string `xml:",chardata"`
			Hint string `xml:"hint,attr"`
		} `xml:"rFonts"`
	} `xml:"rPr"`
}

// Numbering is
type Numbering struct {
	XMLName     xml.Name `xml:"numbering"`
	Text        string   `xml:",chardata"`
	Wpc         string   `xml:"wpc,attr"`
	Cx          string   `xml:"cx,attr"`
	Cx1         string   `xml:"cx1,attr"`
	Mc          string   `xml:"mc,attr"`
	O           string   `xml:"o,attr"`
	R           string   `xml:"r,attr"`
	M           string   `xml:"m,attr"`
	V           string   `xml:"v,attr"`
	Wp14        string   `xml:"wp14,attr"`
	Wp          string   `xml:"wp,attr"`
	W10         string   `xml:"w10,attr"`
	W           string   `xml:"w,attr"`
	W14         string   `xml:"w14,attr"`
	W15         string   `xml:"w15,attr"`
	W16se       string   `xml:"w16se,attr"`
	Wpg         string   `xml:"wpg,attr"`
	Wpi         string   `xml:"wpi,attr"`
	Wne         string   `xml:"wne,attr"`
	Wps         string   `xml:"wps,attr"`
	Ignorable   string   `xml:"Ignorable,attr"`
	AbstractNum []struct {
		Text                       string         `xml:",chardata"`
		AbstractNumID              string         `xml:"abstractNumId,attr"`
		RestartNumberingAfterBreak string         `xml:"restartNumberingAfterBreak,attr"`
		Nsid                       TextVal        `xml:"nsid"`
		MultiLevelType             TextVal        `xml:"multiLevelType"`
		Tmpl                       TextVal        `xml:"tmpl"`
		Lvl                        []NumberingLvl `xml:"lvl"`
	} `xml:"abstractNum"`
	Num []struct {
		Text          string  `xml:",chardata"`
		NumID         string  `xml:"numId,attr"`
		AbstractNumID TextVal `xml:"abstractNumId"`
	} `xml:"num"`
}

// Config holds conversion options.
type Config struct {
	Embed     bool
	HTMLTable bool
}

type file struct {
	rels Relationships
	num  Numbering
	r    *zip.ReadCloser
	cfg  Config
	list map[string]int
}

// Node is
type Node struct {
	XMLName xml.Name
	Attrs   []xml.Attr `xml:"-"`
	Content []byte     `xml:",innerxml"`
	Nodes   []Node     `xml:",any"`
}

// UnmarshalXML is
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
		if zf.cfg.Embed {
			fmt.Fprintf(w, "![](data:image/png;base64,%s)",
				base64.StdEncoding.EncodeToString(b[:n]))
		} else {
			err = os.WriteFile(rel.Target, b, 0644)
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
		hasNumPr := false
		for _, n := range node.Nodes {
			if n.XMLName.Local == "numPr" {
				hasNumPr = true
				break
			}
		}
		for _, n := range node.Nodes {
			switch n.XMLName.Local {
			case "ind":
				if !hasNumPr {
					if left, ok := attr(n.Attrs, "left"); ok {
						if i, err := strconv.Atoi(left); err == nil && i > 0 {
							fmt.Fprint(w, strings.Repeat("  ", i/360))
						}
					}
				}
			case "pStyle":
				if val, ok := attr(n.Attrs, "val"); ok {
					if strings.HasPrefix(val, "Heading") {
						if i, err := strconv.Atoi(val[7:]); err == nil && i > 0 {
							fmt.Fprint(w, "\n"+strings.Repeat("#", i)+" ")
						}
					} else if val == "Code" {
						code = true
					} else {
						if i, err := strconv.Atoi(val); err == nil && i > 0 {
							fmt.Fprint(w, "\n"+strings.Repeat("#", i)+" ")
						}
					}
				}
			case "numPr":
				numID := ""
				ilvl := ""
				ilvlNum := 0
				numFmt := ""
				start := 1
				for _, nn := range n.Nodes {
					if nn.XMLName.Local == "numId" {
						if val, ok := attr(nn.Attrs, "val"); ok {
							numID = val
						}
					}
					if nn.XMLName.Local == "ilvl" {
						if val, ok := attr(nn.Attrs, "val"); ok {
							ilvl = val
							if i, err := strconv.Atoi(val); err == nil {
								ilvlNum = i
							}
						}
					}
				}
				for _, num := range zf.num.Num {
					if numID != num.NumID {
						continue
					}
					for _, abnum := range zf.num.AbstractNum {
						if abnum.AbstractNumID != num.AbstractNumID.Val {
							continue
						}
						for _, ablvl := range abnum.Lvl {
							if ablvl.Ilvl != ilvl {
								continue
							}
							if i, err := strconv.Atoi(ablvl.Start.Val); err == nil {
								start = i
							}
							numFmt = ablvl.NumFmt.Val
							break
						}
						break
					}
					break
				}

				key := fmt.Sprintf("%s:%d", numID, ilvlNum)
				if _, ok := zf.list[key]; !ok {
					fmt.Fprint(w, "\n")
				}
				fmt.Fprint(w, strings.Repeat("  ", ilvlNum))
				switch numFmt {
				case "decimal", "aiueoFullWidth":
					cur, ok := zf.list[key]
					if !ok {
						zf.list[key] = start
					} else {
						zf.list[key] = cur + 1
					}
					fmt.Fprintf(w, "%d. ", zf.list[key])
				case "bullet":
					if _, ok := zf.list[key]; !ok {
						zf.list[key] = 1
					}
					fmt.Fprint(w, "* ")
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
		fmt.Fprint(w, "\n")

		type cellInfo struct {
			content  string
			gridSpan int
			vMerge   string // "restart", "continue", or ""
		}

		var cellRows [][]cellInfo
		for _, tr := range node.Nodes {
			if tr.XMLName.Local != "tr" {
				continue
			}
			var cells []cellInfo
			for _, tc := range tr.Nodes {
				if tc.XMLName.Local != "tc" {
					continue
				}
				ci := cellInfo{gridSpan: 1}
				for _, n := range tc.Nodes {
					if n.XMLName.Local != "tcPr" {
						continue
					}
					for _, nn := range n.Nodes {
						switch nn.XMLName.Local {
						case "gridSpan":
							if val, ok := attr(nn.Attrs, "val"); ok {
								if v, err := strconv.Atoi(val); err == nil {
									ci.gridSpan = v
								}
							}
						case "vMerge":
							if val, ok := attr(nn.Attrs, "val"); ok {
								ci.vMerge = val
							} else {
								ci.vMerge = "continue"
							}
						}
					}
				}
				var cbuf bytes.Buffer
				if err := zf.walk(&tc, &cbuf); err != nil {
					return err
				}
				ci.content = strings.Replace(cbuf.String(), "\n", "", -1)
				cells = append(cells, ci)
			}
			cellRows = append(cellRows, cells)
		}

		// Check if table has any merged cells
		hasMerge := false
		for _, cells := range cellRows {
			for _, ci := range cells {
				if ci.gridSpan > 1 || ci.vMerge != "" {
					hasMerge = true
					break
				}
			}
			if hasMerge {
				break
			}
		}

		if hasMerge && zf.cfg.HTMLTable {
			// Calculate rowspan for vMerge cells
			type htmlCell struct {
				content string
				colspan int
				rowspan int
				skip    bool
			}
			htmlRows := make([][]htmlCell, len(cellRows))
			for i, cells := range cellRows {
				htmlRows[i] = make([]htmlCell, len(cells))
				for j, ci := range cells {
					htmlRows[i][j] = htmlCell{
						content: ci.content,
						colspan: ci.gridSpan,
						rowspan: 1,
					}
					if ci.vMerge == "continue" {
						htmlRows[i][j].skip = true
						// Find the restart cell above and increment its rowspan
						for k := i - 1; k >= 0; k-- {
							if j < len(htmlRows[k]) && !htmlRows[k][j].skip {
								htmlRows[k][j].rowspan++
								break
							}
						}
					}
				}
			}

			fmt.Fprint(w, "<table>\n")
			for i, row := range htmlRows {
				fmt.Fprint(w, "  <tr>\n")
				tag := "td"
				if i == 0 {
					tag = "th"
				}
				for _, cell := range row {
					if cell.skip {
						continue
					}
					fmt.Fprintf(w, "    <%s", tag)
					if cell.colspan > 1 {
						fmt.Fprintf(w, " colspan=\"%d\"", cell.colspan)
					}
					if cell.rowspan > 1 {
						fmt.Fprintf(w, " rowspan=\"%d\"", cell.rowspan)
					}
					fmt.Fprintf(w, ">%s</%s>\n", cell.content, tag)
				}
				fmt.Fprint(w, "  </tr>\n")
			}
			fmt.Fprint(w, "</table>\n")
		} else {
			// Plain markdown table (no merged cells)
			var rows [][]string
			for _, cells := range cellRows {
				var cols []string
				for _, ci := range cells {
					cols = append(cols, ci.content)
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
		}
		fmt.Fprint(w, "\n")
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
		var pbuf bytes.Buffer
		for _, n := range node.Nodes {
			if err := zf.walk(&n, &pbuf); err != nil {
				return err
			}
		}
		content := pbuf.String()
		if strings.TrimSpace(content) != "" {
			fmt.Fprint(w, content)
			fmt.Fprintln(w)
		}
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

	b, _ := io.ReadAll(rc)
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
		if ok, _ := path.Match(target, f.Name); ok {
			return f
		}
	}
	return nil
}

func docx2md(arg string, cfg Config) error {
	r, err := zip.OpenReader(arg)
	if err != nil {
		return err
	}
	defer r.Close()

	var rels Relationships
	var num Numbering

	for _, f := range r.File {
		switch f.Name {
		case "word/_rels/document.xml.rels", "word/_rels/document2.xml.rels":
			rc, err := f.Open()
			defer rc.Close()

			b, _ := io.ReadAll(rc)
			if err != nil {
				return err
			}

			err = xml.Unmarshal(b, &rels)
			if err != nil {
				return err
			}
		case "word/numbering.xml":
			rc, err := f.Open()
			defer rc.Close()

			b, _ := io.ReadAll(rc)
			if err != nil {
				return err
			}

			err = xml.Unmarshal(b, &num)
			if err != nil {
				return err
			}
		}
	}

	f := findFile(r.File, "word/document*.xml")
	if f == nil {
		return errors.New("incorrect document")
	}
	node, err := readFile(f)
	if err != nil {
		return err
	}

	var buf bytes.Buffer
	zf := &file{
		r:    r,
		rels: rels,
		num:  num,
		cfg:  cfg,
		list: make(map[string]int),
	}
	err = zf.walk(node, &buf)
	if err != nil {
		return err
	}
	fmt.Print(buf.String())

	return nil
}

func main() {
	var cfg Config
	var showVersion bool
	flag.BoolVar(&cfg.Embed, "embed", false, "embed resources")
	flag.BoolVar(&cfg.HTMLTable, "html-table", false, "output merged cells as HTML table")
	flag.BoolVar(&showVersion, "v", false, "Print the version")
	flag.Parse()
	if showVersion {
		fmt.Printf("%s %s (rev: %s/%s)\n", name, version, revision, runtime.Version())
		return
	}
	if flag.NArg() == 0 {
		flag.Usage()
		os.Exit(1)
	}
	for _, arg := range flag.Args() {
		if err := docx2md(arg, cfg); err != nil {
			log.Fatal(err)
		}
	}
}
