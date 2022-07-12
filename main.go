package main

import (
	"fmt"
	"github.com/unidoc/unioffice/common/license"
	"github.com/unidoc/unioffice/document"
	"log"
)

func init() {
	key := "385f0de7ecde7d95f9254dc2e5b61d206978cbb6f301fcd75bf2acc3bb1290b5"

	if err := license.SetMeteredKey(key); err != nil {
		panic(err)
	}
}

func Products() []Product {
	p1 := Product{
		Name:     "Моцарелла",
		Price:    399,
		Quantity: 1,
	}
	p2 := Product{
		Name:     "Грибная",
		Price:    417,
		Quantity: 2,
	}
	products := []Product{p1, p2}
	return products
}

type Product struct {
	Name     string
	Price    int64
	Quantity uint64
}

func main() {

	doc, err := document.Open("templ.docx")
	if err != nil {
		log.Fatalf("error opening Windows Word 2016 document: %s", err)
	}
	defer doc.Close()

	var paragraphs []document.Paragraph
	for _, p := range doc.Paragraphs() {
		paragraphs = append(paragraphs, p)
	}
	//Filling header placeholders
	hh := doc.Headers()
	for _, h := range hh {
		for _, p := range h.Paragraphs() {
			for _, r := range p.Runs() {
				switch r.Text() {
				case "ord":
					r.ClearContent()
					r.AddText("#000012")
				}
			}
		}
	}
	//Filling footer placeholders
	ff := doc.Footers()
	for _, f := range ff {
		for _, p := range f.Paragraphs() {
			for _, r := range p.Runs() {
				switch r.Text() {
				case "sum":
					r.ClearContent()
					r.AddText("1854.0₽")
				}

			}
		}
	}
	//Filling body placeholders
	for _, p := range paragraphs {
		for _, r := range p.Runs() {
			switch r.Text() {
			case "username":
				r.ClearContent()
				r.AddText("Артем Тимофеев")
			case "phoneNumber":
				r.ClearContent()
				r.AddText("+79524000770")
			case "address":
				r.ClearContent()
				r.AddText("Пушкинская 29а")
			case "delivery":
				r.ClearContent()
				r.AddText("Да")
			case "ent":
				r.ClearContent()
				r.AddText("3")
			case "fl":
				r.ClearContent()
				r.AddText("5")
			case "gr":
				r.ClearContent()
				r.AddText("3")
			case "Содержимое":
				r.ClearContent()
				r.AddText("Содержимое:")
				r.AddBreak()
				pp := Products()
				for _, p := range pp {
					r.AddText(" - ")
					r.AddText(p.Name)
					r.AddTab()
					r.AddTab()
					r.AddText(fmt.Sprintf("%d * %d.0₽", p.Quantity, p.Price))
					r.AddBreak()
				}
			}

		}
	}
	if err = doc.SaveToFile("edit-document.docx"); err != nil {
		panic(err)
	}
}
