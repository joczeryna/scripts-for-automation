package main

import (
	"context"
	"encoding/json"
	"fmt"
	"log"
	"os"
	"phonelog/internal/rwsheets"
	"strconv"

	"github.com/joho/godotenv"
	sheets "google.golang.org/api/sheets/v4"
)

func main() {
	// Load the environment file
	if err := godotenv.Load(); err != nil {
		log.Fatalf("error loading env file: %v", err)
	}

	filename := os.Getenv("FILE_PATH")
	if filename == "" {
		log.Fatalf("invalid filepath given for data file")
	}
	log.Print(filename)

	ctx := context.Background()

	sheetsToken := os.Getenv("TOKEN")
	sheetsCred := os.Getenv("CREDENTIALS")
	sheetsScope := "https://www.googleapis.com/auth/spreadsheets"

	if sheetsToken == "" {
		log.Fatalf("invalid token path provided")
	}

	if sheetsCred == "" {
		log.Fatalf("invalid credential file path provided")
	}

	srv, err := rwsheets.NewSheetsService(ctx, sheetsCred, sheetsToken, sheetsScope)
	if err != nil {
		log.Fatalf("failed to create new sheets service: %s", err.Error())
	}

	data, err := getSmartsheetData(filename)
	if err != nil {
		log.Fatal(err)
	}

	if err := updatePhoneLog(data, srv); err != nil {
		log.Fatalf("failed to update phone log - %s", err.Error())
	}

	if err := updateEEData(data, srv); err != nil {
		log.Fatalf("failed to update EE Data sheet - %s", err.Error())
	}
}

type EmployeeReport struct {
	Data map[string]*EmployeeData `json:"data"`
}

type EmployeeData struct {
	WorkerNo        string `json:"worker_no"`
	Name            string `json:"name"`
	Title           string `json:"title"`
	City            string `json:"city"`
	State           string `json:"state"`
	DivisionManager string `json:"division_manager"`
	OfficeNbr       string `json:"office_nbr"`
	Division        string `json:"division"`
	Manager         string `json:"manager"`
	Supervisor      string `json:"supervisor"`
	Status          string `json:"status"`
	WorkType        string `json:"work_type"`
	Pv              string `json:"pvi"`
	HireDate        string `json:"hire_date"`
	TermDate        string `json:"term_date"`
	Email           string `json:"email"`
	Phone           string `json:"phone"`
	Ad              string `json:"ad"`
}

// getSmartsheetData: Reads the Smartsheet data from the given file and returns it
func getSmartsheetData(filename string) (EmployeeReport, error) {
	var report EmployeeReport

	// Read the data from the file
	data, err := os.ReadFile(filename)
	if err != nil {
		return report, fmt.Errorf("failed to read data from file %s: %s", filename, err.Error())
	}

	// Unmarshal the data from the file into our Go struct
	if err := json.Unmarshal(data, &report); err != nil {
		return report, fmt.Errorf("failed to unmarshal data from file %s: %s", filename, err.Error())
	}

	return report, nil
}

// getRowCount: Reads and parses the row the row count from the spreadsheet.
func getRowCount(ssid, readRange string, srv *sheets.Service) (int, error) {
	rows, err := rwsheets.GetSheetData(ssid, readRange, srv)
	if err != nil {
		return -1, err
	}
	row := rows[0].Values
	countStr := row[0].FormattedValue
	count, err := strconv.Atoi(countStr)
	if err != nil {
		return -1, fmt.Errorf("failed to parse row count %s: %s", countStr, err.Error())
	}

	return count, nil
}

// emptyRow: Returns an array of CellData (a sheets Row) with empty values
// for the given number of cells
func emptyRow(numCells int) []*sheets.CellData {
	var values []*sheets.CellData
	for x := 0; x < numCells; x++ {
		cell := sheets.CellData{
			UserEnteredValue: rwsheets.TextValue(""),
		}
		values = append(values, &cell)
	}
	return values
}

func updatePhoneLog(data EmployeeReport, srv *sheets.Service) error {
	ssid := os.Getenv("PHONE_LOG_SSID")
	gidStr := os.Getenv("PHONE_LOG_GID")
	gid, err := strconv.ParseInt(gidStr, 10, 64)
	if err != nil {
		return fmt.Errorf("failed to parse sheet ID %s from environment variable: %s", gidStr, err.Error())
	}

	countRange := os.Getenv("PHONE_LOG_COUNT_READ_RANGE")
	count, err := getRowCount(ssid, countRange, srv)
	if err != nil {
		return err
	}

	if count == 0 {
		// We have no data in our report, let's create from scratch
		return createPhoneLogFromScratch(data, gid, ssid, srv)
	}

	return updatePhoneLogData(data, count, gid, ssid, srv)
}

func updatePhoneLogData(report EmployeeReport, count int, gid int64, ssid string, srv *sheets.Service) error {
	styler := rwsheets.NewStyler().FontBold(false).FontFamily("Calibri").FontSize(int64(11)).HorizontalAlignment("LEFT").VerticalAlignment("MIDDLE")

	var newRows []*sheets.RowData
	for _, ee := range report.Data {
		if ee.Status != "Active" {
			continue
		}

		if ee.WorkerNo == "073008" || ee.WorkerNo == "088665" || ee.WorkerNo == "824621" || ee.WorkerNo == "072878" {
			// Not including exectuive leadership
			continue
		}

		cells := []*sheets.CellData{
			styler.TextCell(ee.WorkerNo, nil),
			styler.TextCell(ee.Name, nil),
			styler.TextCell(ee.Title, nil),
			styler.TextCell(ee.City, nil),
			styler.TextCell(ee.State, nil),
			styler.TextCell(ee.DivisionManager, nil),
			styler.TextCell(ee.Email, nil),
			styler.TextCell(ee.Phone, nil),
		}
		row := sheets.RowData{
			Values: cells,
		}
		newRows = append(newRows, &row)
	}

	rowCount := len(newRows)
	if rowCount < count {
		for i := rowCount; i <= count; i++ {
			newRow := sheets.RowData{
				Values: emptyRow(8),
			}
			newRows = append(newRows, &newRow)
		}
	}

	// Update our sheet with the new row data
	if err := rwsheets.UpdateSheetData(ssid, int64(8), gid, int64(0), int64(1), newRows, srv); err != nil {
		return err
	}

	return nil
}

func createPhoneLogFromScratch(report EmployeeReport, gid int64, ssid string, srv *sheets.Service) error {
	styler := rwsheets.NewStyler().FontBold(true).FontFamily("Calibri").FontSize(int64(11)).HorizontalAlignment("LEFT").VerticalAlignment("MIDDLE")
	headerValues := []string{
		"Worker No",
		"Name",
		"Title",
		"Reporting Office",
		"State",
		"Division/Manager",
		"Email Address",
		"Phone",
	}

	// Create new header rows for the data
	newRows := styler.CreateHeaderRow(headerValues, nil)

	// Set the stlyer fontbold to false (we only want the header to be bold)
	styler = styler.FontBold(false)

	for _, ee := range report.Data {
		if ee.Status != "Active" {
			continue
		}

		if ee.WorkerNo == "073008" || ee.WorkerNo == "088665" || ee.WorkerNo == "824621" || ee.WorkerNo == "072878" {
			// Not including exectuive leadership
			continue
		}

		cells := []*sheets.CellData{
			styler.TextCell(ee.WorkerNo, nil),
			styler.TextCell(ee.Name, nil),
			styler.TextCell(ee.Title, nil),
			styler.TextCell(ee.City, nil),
			styler.TextCell(ee.State, nil),
			styler.TextCell(ee.DivisionManager, nil),
			styler.TextCell(ee.Email, nil),
			styler.TextCell(ee.Phone, nil),
		}
		row := sheets.RowData{
			Values: cells,
		}
		newRows = append(newRows, &row)
	}

	if err := rwsheets.UpdateSheetData(ssid, int64(8), gid, int64(0), int64(0), newRows, srv); err != nil {
		return err
	}

	return nil
}

func updateEEData(data EmployeeReport, srv *sheets.Service) error {
	ssid := os.Getenv("EE_DATA_SSID")
	gidStr := os.Getenv("EE_DATA_GID")
	gid, err := strconv.ParseInt(gidStr, 10, 64)
	if err != nil {
		return fmt.Errorf("failed to parse sheet ID %s from environment variable: %s", gidStr, err.Error())
	}

	countRange := os.Getenv("EE_DATA_COUNT_READ_RANGE")
	count, err := getRowCount(ssid, countRange, srv)
	if err != nil {
		return err
	}

	if count == 0 {
		// We have no data in our report, let's create from scratch
		return createEmployeeDataFromScratch(data, gid, ssid, srv)
	}

	readRangePartial := os.Getenv("EE_DATA_READ_RANGE")
	readRange := fmt.Sprintf("%s%d", readRangePartial, count)

	return updateEmployeeData(data, count, gid, readRange, ssid, srv)
}

func createEmployeeDataFromScratch(report EmployeeReport, gid int64, ssid string, srv *sheets.Service) error {
	styler := rwsheets.NewStyler().FontBold(true).FontFamily("Calibri").FontSize(int64(11)).HorizontalAlignment("LEFT").VerticalAlignment("MIDDLE")

	headerValues := []string{
		"Worker No",
		"Name",
		"Title",
		"City",
		"State",
		"Division/Manager",
		"Office Nbr",
		"Division",
		"Manager",
		"Supervisor",
		"Status",
		"Work Type",
		"PVI",
		"Hire Date",
		"Term Date",
		"Email",
		"Phone",
		"Employee AD",
	}

	// Create new header rows for the data
	newRows := styler.CreateHeaderRow(headerValues, nil)

	// Set the stlyer fontbold to false (we only want the header to be bold)
	styler = styler.FontBold(false)

	for _, ee := range report.Data {
		cells := []*sheets.CellData{
			styler.TextCell(ee.WorkerNo, nil),
			styler.TextCell(ee.Name, nil),
			styler.TextCell(ee.Title, nil),
			styler.TextCell(ee.City, nil),
			styler.TextCell(ee.State, nil),
			styler.TextCell(ee.OfficeNbr, nil),
			styler.TextCell(ee.DivisionManager, nil),
			styler.TextCell(ee.Division, nil),
			styler.TextCell(ee.Manager, nil),
			styler.TextCell("", nil),
			styler.TextCell(ee.Status, nil),
			styler.TextCell(ee.WorkType, nil),
			styler.TextCell(ee.Pv, nil),
			styler.TextCell(ee.HireDate, nil),
			styler.TextCell(ee.TermDate, nil),
			styler.TextCell(ee.Email, nil),
			styler.TextCell(ee.Phone, nil),
			styler.TextCell(ee.Ad, nil),
		}
		row := sheets.RowData{
			Values: cells,
		}
		newRows = append(newRows, &row)
	}

	if err := rwsheets.UpdateSheetData(ssid, int64(18), gid, int64(0), int64(1), newRows, srv); err != nil {
		return err
	}

	return nil
}

func updateEmployeeData(report EmployeeReport, count int, gid int64, readRange, ssid string, srv *sheets.Service) error {
	styler := rwsheets.NewStyler().FontBold(false).FontFamily("Calibri").FontSize(int64(11)).HorizontalAlignment("LEFT").VerticalAlignment("MIDDLE")

	currRows, err := rwsheets.GetSheetData(ssid, readRange, srv)
	if err != nil {
		return err
	}

	for i, row := range currRows {
		if i == 0 {
			// Ignoring header row
			continue
		}

		values := row.Values
		workerNo := values[0].FormattedValue

		ee, ok := report.Data[workerNo]
		if !ok {
			log.Printf("WARNING: Employee %s listed on sheet is not in Smartsheet data", workerNo)
			continue
		}
		values = updateEmployeeRowData(ee, values, styler)
		row.Values = values
		currRows[i] = row
		delete(report.Data, workerNo)
	}

	if len(report.Data) > 0 {
		currRows = makeEmployeeRows(currRows, report, styler)
	}

	if err := rwsheets.UpdateSheetData(ssid, int64(18), gid, int64(0), int64(1), currRows, srv); err != nil {
		return err
	}

	return nil
}

func updateEmployeeRowData(ee *EmployeeData, cells []*sheets.CellData, styler *rwsheets.Styler) []*sheets.CellData {
	if len(cells) < 18 {
		return cells
	}

	nameIdx := 1
	titleIdx := 2
	cityIdx := 3
	stateIdx := 4
	divMgrIdx := 5
	officeIdx := 6
	divisionIdx := 7
	managerIdx := 8
	statusIdx := 10
	typeIdx := 11
	pvIdx := 12
	hireIdx := 13
	termIdx := 14
	emailIdx := 15
	phoneIdx := 16
	adIdx := 17

	name := cells[nameIdx].FormattedValue
	if name != ee.Name {
		cells[nameIdx] = styler.FontColor(rwsheets.BLUE_COLOR).TextCell(ee.Name, nil)
	}
	title := cells[titleIdx].FormattedValue
	if title != ee.Title {
		cells[titleIdx] = styler.FontColor(rwsheets.BLUE_COLOR).TextCell(ee.Title, nil)
	}
	city := cells[cityIdx].FormattedValue
	if city != ee.City {
		cells[cityIdx] = styler.FontColor(rwsheets.BLUE_COLOR).TextCell(ee.City, nil)
	}
	state := cells[stateIdx].FormattedValue
	if state != ee.State {
		cells[stateIdx] = styler.FontColor(rwsheets.BLUE_COLOR).TextCell(ee.State, nil)
	}
	divMgr := cells[divMgrIdx].FormattedValue
	if divMgr != ee.DivisionManager {
		cells[divMgrIdx] = styler.FontColor(rwsheets.BLUE_COLOR).TextCell(ee.DivisionManager, nil)
	}
	officeNbr := cells[officeIdx].FormattedValue
	if officeNbr != ee.OfficeNbr {
		cells[officeIdx] = styler.FontColor(rwsheets.BLUE_COLOR).TextCell(ee.OfficeNbr, nil)
	}
	division := cells[divisionIdx].FormattedValue
	if division != ee.Division && ee.Division != "" {
		cells[divisionIdx] = styler.FontColor(rwsheets.BLUE_COLOR).TextCell(ee.Division, nil)
	}
	mgr := cells[managerIdx].FormattedValue
	if mgr != ee.Manager && ee.Manager != "" {
		cells[managerIdx] = styler.FontColor(rwsheets.BLUE_COLOR).TextCell(ee.Manager, nil)
	}
	status := cells[statusIdx].FormattedValue
	if status != ee.Status {
		cells[statusIdx] = styler.FontColor(rwsheets.BLUE_COLOR).TextCell(ee.Status, nil)
	}
	workType := cells[typeIdx].FormattedValue
	if workType != ee.WorkType {
		cells[typeIdx] = styler.FontColor(rwsheets.BLUE_COLOR).TextCell(ee.WorkType, nil)
	}
	pvi := cells[pvIdx].FormattedValue
	if pvi != ee.Pv {
		cells[pvIdx] = styler.FontColor(rwsheets.BLUE_COLOR).TextCell(ee.Pv, nil)
	}
	hireDate := cells[hireIdx].FormattedValue
	if hireDate != ee.HireDate {
		cells[hireIdx] = styler.FontColor(rwsheets.BLUE_COLOR).TextCell(ee.HireDate, nil)
	}
	termDate := cells[termIdx].FormattedValue
	if termDate != ee.TermDate {
		cells[termIdx] = styler.FontColor(rwsheets.BLUE_COLOR).TextCell(ee.TermDate, nil)
	}
	email := cells[emailIdx].FormattedValue
	if email != ee.Email {
		cells[emailIdx] = styler.FontColor(rwsheets.BLUE_COLOR).TextCell(ee.Email, nil)
	}
	phone := cells[phoneIdx].FormattedValue
	if phone != ee.Phone {
		cells[phoneIdx] = styler.FontColor(rwsheets.BLUE_COLOR).TextCell(ee.Phone, nil)
	}
	ad := cells[adIdx].FormattedValue
	if ad != ee.Ad {
		cells[adIdx] = styler.FontColor(rwsheets.BLUE_COLOR).TextCell(ee.Ad, nil)
	}
	return cells
}

func makeEmployeeRows(currRows []*sheets.RowData, report EmployeeReport, styler *rwsheets.Styler) []*sheets.RowData {
	for _, ee := range report.Data {
		cells := []*sheets.CellData{
			styler.TextCell(ee.WorkerNo, nil),
			styler.TextCell(ee.Name, nil),
			styler.TextCell(ee.Title, nil),
			styler.TextCell(ee.City, nil),
			styler.TextCell(ee.State, nil),
			styler.TextCell(ee.DivisionManager, nil),
			styler.TextCell(ee.OfficeNbr, nil),
			styler.TextCell(ee.Division, nil),
			styler.TextCell(ee.Manager, nil),
			styler.TextCell("", nil),
			styler.TextCell(ee.Status, nil),
			styler.TextCell(ee.WorkType, nil),
			styler.TextCell(ee.Pv, nil),
			styler.TextCell(ee.HireDate, nil),
			styler.TextCell(ee.TermDate, nil),
			styler.TextCell(ee.Email, nil),
			styler.TextCell(ee.Phone, nil),
			styler.TextCell(ee.Ad, nil),
		}
		row := sheets.RowData{
			Values: cells,
		}
		currRows = append(currRows, &row)
	}
	return currRows
}
