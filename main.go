package main

import (
	"encoding/csv"
	"fmt"
	"io"
	"log"
	"os"
	"path/filepath"
	"slices"
	"sort"
	"strconv"

	"github.com/gomutex/godocx"
	"github.com/gomutex/godocx/docx"
	"github.com/gomutex/godocx/wml/stypes"
)

type CSVData struct {
	CatalogName string
	Title       string
	Ref         string
	BuyerName   string
	BuyerPseudo string
	BuyerEmail  string
	Address     Address
}

type Address struct {
	Street  string
	ZipCode string
	City    string
	Country string
}

type ClientOrders struct {
	Pseudo  string
	Name    string
	Email   string
	Address Address
	Orders  []Order
}

type Order struct {
	Ref   int
	Title string
}

// getCSVData
// 0 Catalogue ex data: "573e Vente sur Offres"
// 1 Référence ex data: 3472
// 2 Titre ex data: "Lettre. Nouvelle-Zélande. ""Pigeongram Services"". Document privé de la Poste par Pigeon, cachet ovale bleu 18 janv 1903."
// 3 Date ex data: "2025-06-03 12:56:27"
// 4 Montant ex data: 1300
// 5 Devise ex data: EUR
// 6 Acheteur ex data: jstimpy
// 7 Email ex data: Jeanettetesting123@gmail.com
// 8 Coordonnées ex data: "Jeanette Kelly 521b Achilles Avenue Whangamata 3643 Whangamata  NZ +64 21 225 2626"
// 9 Nom ex data: "Jeanette Kelly"
// 10 Adresse ex data: "521b Achilles Avenue Whangamata"
// 11 Code postal ex data: 3643
// 12 Ville ex data: Whangamata
// 13 Région ex data: US-NY
// 14 Pays ex data: NZ
func getCSVData(fileName string) []CSVData {
	f, err := os.Open(fileName)
	if err != nil {
		panic(err)
	}

	r := csv.NewReader(f)
	r.Comma = ';'

	var data []CSVData

	for {
		record, err := r.Read()
		if err == io.EOF {
			break
		}
		if err != nil {
			log.Fatal(err)
		}

		var d CSVData
		d.CatalogName = record[0]
		d.Title = record[4]
		d.Ref = record[1]
		d.BuyerName = record[9]
		d.BuyerPseudo = record[6]
		d.BuyerEmail = record[7]
		d.Address = Address{
			Street:  record[10],
			ZipCode: record[11],
			City:    record[12],
			Country: returnCountry(record[14], record[13]),
		}

		data = append(data, d)
	}
	return data
}

func returnCountry(cc, cn string) string {
	if len(cc) == 0 {
		return cn
	}
	if len(cn) == 0 {
		return cc
	}
	return cc + " - " + cn
}

func setMapOfClientOrders(data []CSVData) map[string]*ClientOrders {
	dataOrders := make(map[string]*ClientOrders)
	for i := range data {
		userKey := data[i].BuyerName + data[i].BuyerEmail
		_, ok := dataOrders[userKey]

		if !ok {
			dataOrders[userKey] = &ClientOrders{
				Name:    data[i].BuyerName,
				Email:   data[i].BuyerEmail,
				Pseudo:  data[i].BuyerPseudo,
				Address: data[i].Address,
			}
		}

		if len(data[i].Ref) > 0 {
			refInt, err := strconv.Atoi(data[i].Ref)
			if err != nil {
				continue
			}

			dataOrders[userKey].Orders = append(dataOrders[userKey].Orders, Order{
				Title: data[i].Title,
				Ref:   refInt,
			})
		}
	}
	return dataOrders
}

func createWord(catalogName string, data map[string]*ClientOrders) *docx.RootDoc {
	// fmt.Println("createWord:", catalogName, data)
	// Create New Document
	document, err := godocx.NewDocument()
	if err != nil {
		log.Fatal(err)
	}

	var listKeys []string
	for k := range data {
		listKeys = append(listKeys, k)
	}

	slices.Sort(listKeys)

	for i := 0; i < len(listKeys); i++ {
		d := data[listKeys[i]]
		document.AddHeading(catalogName, 0)
		namep := document.AddParagraph("")
		namep.AddText(d.Name).Bold(true).Size(17)
		addressp := document.AddParagraph("")
		lineBreak := stypes.BreakTypeTextWrapping
		addressp.AddText(d.Address.Street).Size(15).AddBreak(&lineBreak)
		addressp.AddText(d.Address.ZipCode + " " + d.Address.City).Size(15).AddBreak(&lineBreak)
		addressp.AddText(d.Address.Country).Size(15).AddBreak(&lineBreak)
		addressp.AddText("Pseudo :  " + d.Pseudo).Size(15).AddBreak(&lineBreak)
		addressp.AddText("Email :  " + d.Email).Size(15).AddBreak(&lineBreak)
		addressp.AddText("").Size(15).AddBreak(&lineBreak)

		if len(d.Orders) > 1 {
			sort.Slice(d.Orders, func(i, j int) bool { return d.Orders[i].Ref < d.Orders[j].Ref })
		}

		for i := 0; i < len(d.Orders); i++ {
			p := document.AddParagraph("") // paragraphe vide
			ref := fmt.Sprint(d.Orders[i].Ref)
			run := p.AddText("Lot " + ref + " Offre " + d.Orders[i].Title) // ajoute le texte
			run.Bold(true).Size(15).AddBreak(&lineBreak)
		}
		document.AddPageBreak()
	}

	return document
}

func main() {
	exePath, err := os.Executable()
	if err != nil {
		log.Fatal(err)
	}
	baseDir := filepath.Dir(exePath)

	csvPath := filepath.Join(baseDir, "input.csv")
	outPath := filepath.Join(baseDir, "output.docx")

	// dataCSV := getCSVData("testdata/1-catalog_10993.csv")
	dataCSV := getCSVData(csvPath)

	dataOrders := setMapOfClientOrders(dataCSV)
	document := createWord(dataCSV[1].CatalogName, dataOrders)
	// Save the modified document to a new file
	// err := document.SaveTo("demo.docx")
	err = document.SaveTo(outPath)
	if err != nil {
		log.Fatal(err)
	}
}
