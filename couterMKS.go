package main

import (
	"fmt"
	"net/url"
	"strconv"
	"strings"
	"time"

	"github.com/gocolly/colly/v2"
	"github.com/xuri/excelize/v2"
)

type couterMSK_data struct {
	delo          string // Номер дела
	matrial       string // Номер материала
	link          string // Ссылка на дело
	istec         string // Истец
	otvetchik     string // Ответчик
	status        string // Текущее состояние
	couter_people string // Судья
	statya        string // Статья
	category      string // Категория дела
}

const url_couter string = "https://mos-gorsud.ru"
const url_MSK string = "/search?caseDateFrom=DATEFROM&caseDateTo=DATETO&caseFinalDateFrom=&caseFinalDateTo=&caseLegalForceDateFrom=&caseLegalForceDateTo=&caseNumber=&letterNumber=&category=CATEGORY&codex=&courtAlias=&docsDateFrom=&docsDateTo=&documentStatus=&documentText=&documentType=&instance=1&judge=&participant=&processType=2&publishingState=&uid=&year=&formType=fullForm&page=PAGE"

var index_global int = 2

func main() {
	var categorys = []string{"497432a", "2729f8c", "adbb14a", "4dd038b", "2eb1caa", "078ed6e", "ad74a31", "a261baa", "0d7a99b", "ab1b404", "2a10b93", "6b5319e", "42d09ab", "24413a3", "60e674b", "af7a9a6"}

	var link string
	var page_all, page_tecal int
	var handle couterMSK_data

	dt := time.Now()
	dateTo := dt.Format("02.01.2006")
	dateFrom := dt.AddDate(0, -3, 0).Format("02.01.2006")
	outputfilename := "Суды МСК от " + dateFrom + " до " + dateTo + ".xlsx"

	c := colly.NewCollector()
	fileOutput := excelize.NewFile()
	fileOutput.NewSheet("main")
	fileOutput.DeleteSheet("Sheet1")
	make_Title(fileOutput)
	c.OnHTML("div[class=paginationContainer]", func(element *colly.HTMLElement) { // Получить количество листов
		response, err_page_all := element.DOM.Find("a[class=intheend]").Attr("href")
		if err_page_all {
			u, _ := url.Parse(response)        // всю ссылку
			m, _ := url.ParseQuery(u.RawQuery) // только содержимое
			page_all, _ = strconv.Atoi(m["page"][0])
		} else {
			page_all = 1
		}
	})

	c.OnHTML("div[class=wrapper-search-tables] tbody tr", func(element *colly.HTMLElement) { // Спарсить таблицу данных
		//response := element.DOM.Find("td:nth-child(1)").Text()

		handle.delo = element.DOM.Find("td:nth-child(1) a").Text()
		handle.link, _ = element.DOM.Find("td:nth-child(1) a").Attr("href")
		handle.link = url_couter + handle.link
		handle.matrial = element.DOM.Find("td:nth-child(1) nobr:nth-child(1)").Text()

		handle.istec = element.DOM.Find("td:nth-child(2) div[class=right]").Text()
		handle.istec = strings.Replace(handle.istec, "Истец:", "", 1)
		str_ist_otv := strings.Split(handle.istec, "Ответчик:")
		if len(str_ist_otv) == 2 {
			handle.istec = str_ist_otv[0]
			handle.otvetchik = str_ist_otv[1]
		}

		handle.status = element.DOM.Find("td:nth-child(3)").Text()

		handle.couter_people = element.DOM.Find("td:nth-child(4)").Text()

		handle.statya = element.DOM.Find("td:nth-child(5)").Text()

		handle.category = element.DOM.Find("td:nth-child(6)").Text()

		handle = trimStruct(handle)
		saveTypeOnXLSX(fileOutput, handle)
	})

	for _, tecal_category := range categorys {
		page_all = 1
		for page_tecal = 1; page_tecal <= page_all; page_tecal++ {
			link = url_couter + url_MSK
			link = strings.Replace(link, "DATEFROM", dateFrom, 1)
			link = strings.Replace(link, "DATETO", dateTo, 1)
			link = strings.Replace(link, "CATEGORY", tecal_category, 1)
			link = strings.Replace(link, "PAGE", strconv.Itoa(page_tecal), 1)
			c.Visit(link)
		}
	}
	// fmt.Println("Pages: ", page_all)

	if err := fileOutput.SaveAs(outputfilename); err != nil {
		fmt.Println(err)
	}
}
func make_Title(f *excelize.File) {
	f.SetCellValue("main", "A1", "Номер дела")
	f.SetCellValue("main", "B1", "Номер материала")
	f.SetCellValue("main", "C1", "Ссылка на дело")
	f.SetCellValue("main", "D1", "Истец")
	f.SetCellValue("main", "E1", "Ответчик")
	f.SetCellValue("main", "F1", "Текущее состояние")
	f.SetCellValue("main", "G1", "Судья")
	f.SetCellValue("main", "H1", "Статья")
	f.SetCellValue("main", "I1", "Категория дела")
}
func saveTypeOnXLSX(f *excelize.File, cou couterMSK_data) {
	f.SetCellValue("main", "A"+strconv.Itoa(index_global), cou.delo)
	f.SetCellValue("main", "B"+strconv.Itoa(index_global), cou.matrial)
	f.SetCellValue("main", "C"+strconv.Itoa(index_global), cou.link)
	f.SetCellValue("main", "D"+strconv.Itoa(index_global), cou.istec)
	f.SetCellValue("main", "E"+strconv.Itoa(index_global), cou.otvetchik)
	f.SetCellValue("main", "F"+strconv.Itoa(index_global), cou.status)
	f.SetCellValue("main", "G"+strconv.Itoa(index_global), cou.couter_people)
	f.SetCellValue("main", "H"+strconv.Itoa(index_global), cou.statya)
	f.SetCellValue("main", "I"+strconv.Itoa(index_global), cou.category)
	index_global++
}
func trimAll(str string) string {
	return strings.TrimSpace(str)
}
func trimStruct(data couterMSK_data) couterMSK_data {
	output := couterMSK_data{
		delo:          trimAll(data.delo),
		matrial:       trimAll(data.matrial),
		link:          trimAll(data.link),
		istec:         trimAll(data.istec),
		otvetchik:     trimAll(data.otvetchik),
		status:        trimAll(data.status),
		couter_people: trimAll(data.couter_people),
		statya:        trimAll(data.statya),
		category:      trimAll(data.category),
	}
	return output
}
