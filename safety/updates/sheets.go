package updates

import (
	"fmt"
	"log"
	"strings"
	"time"

	rwsheets "github.com/cryliss/go-rwsheets"
	sheets "google.golang.org/api/sheets/v4"
)

func (u *UpdateManager) updateIssuesSheet(ds map[string]*DeviationSummary) error {
	// Get our sheets service
	srv, err := u.getSheets()
	if err != nil {
		return err
	}

	u.mu.Lock()
	sheet := u.sheetData["issues"]
	u.mu.Unlock()

	rows, err := rwsheets.GetSheetData(sheet.Ssid, sheet.ReadRange, srv)
	if err != nil {
		log.Printf("failed to get sheet %s data - %s", sheet.Name, err.Error())
		return err
	}

	batchUpdate := sheets.BatchUpdateSpreadsheetRequest{
		IncludeSpreadsheetInResponse: false,
	}

	today := time.Now().Format("1/2/2006")

	lastRow := 0
	for i, row := range rows {
		values := row.Values

		auditID := values[1].FormattedValue
		if auditID == "" {
			log.Printf("Presumeably the first row (%d) in the sheet without values - exiting loop", i)
			lastRow = i
			break
		}

		vehicle := values[3].FormattedValue
		qrCode := values[4].FormattedValue
		if len(vehicle) < 5 && (len(qrCode) > 8 || qrCode == "") {
			qrCell := values[4]
			rowNum := i + 1
			qrFormula := fmt.Sprintf(`=XLOOKUP($D%d, Equipment!A:A, Equipment!B:B, "", 0)`, rowNum)
			qrCell.UserEnteredValue = rwsheets.FormulaValue(qrFormula)
			qrReq := u.makeUpdateRequest(u.makeNewCellRow(qrCell), sheet, sheet.IssuesCol-2, int64(i))
			batchUpdate.Requests = append(batchUpdate.Requests, qrReq)
		}

		dev, ok := ds[auditID]
		if !ok {
			log.Printf("Did not find audit ID %s in deviation summary", auditID)
			continue
		}
		dev.Found = true
		ds[auditID] = dev

		issues := values[sheet.IssuesCol].FormattedValue
		if issues != dev.Issues {
			newCell := values[sheet.IssuesCol]
			newCell.UserEnteredValue = rwsheets.TextValue(dev.Issues)
			req := u.makeUpdateRequest(u.makeNewCellRow(newCell), sheet, sheet.IssuesCol, int64(i))
			batchUpdate.Requests = append(batchUpdate.Requests, req)
		}

		if dev.Open == "" {
			status := values[sheet.StatusCol].FormattedValue
			if status != "Closed" {
				newStatusCell := values[sheet.StatusCol]
				newStatusCell.UserEnteredValue = rwsheets.TextValue("Closed")
				reqStatus := u.makeUpdateRequest(u.makeNewCellRow(newStatusCell), sheet, sheet.StatusCol, int64(i))
				batchUpdate.Requests = append(batchUpdate.Requests, reqStatus)

				if sheet.CompCol > int64(len(values)) || values[sheet.CompCol].FormattedValue != "" {
					continue
				}
				newCompCell := values[sheet.CompCol]
				newCompCell.UserEnteredValue = rwsheets.TextValue(today)
				reqComp := u.makeUpdateRequest(u.makeNewCellRow(newCompCell), sheet, sheet.CompCol, int64(i))
				batchUpdate.Requests = append(batchUpdate.Requests, reqComp)
			}
		}

	}

	startRow := lastRow
	var newRows []*sheets.RowData
	for _, dev := range ds {
		if dev.Found || !strings.Contains(dev.Type, "DVIR") {
			continue
		}

		if lastRow > len(rows)-1 {
			break
		}

		row := rows[lastRow]
		lastRow++
		values := row.Values

		for col, cell := range values {
			switch col {
			case 0:
				dateValue, _ := rwsheets.SerialDate(dev.Date, "1/2/2006")
				cell.UserEnteredValue = rwsheets.NumberValue(dateValue)
				values[col] = cell
			case 1:
				cell.UserEnteredValue = rwsheets.FormulaValue(dev.LinkFormula)
				values[col] = cell
			case 2:
				formula := fmt.Sprintf(`=XLOOKUP("%s", Equipment!A:A, Equipment!C:C, "%s", 0)`, dev.AssetNbr, dev.Division)
				cell.UserEnteredValue = rwsheets.FormulaValue(formula)
				values[col] = cell
			case 3:
				cell.UserEnteredValue = rwsheets.TextValue(dev.AssetNbr)
				values[col] = cell
			case 4:
				qrformula := fmt.Sprintf(`=XLOOKUP(D%d, Equipment!A:A, Equipment!B:B, "", 0)`, lastRow)
				cell.UserEnteredValue = rwsheets.FormulaValue(qrformula)
				values[col] = cell
			case 5:
				cell.UserEnteredValue = rwsheets.TextValue(dev.Driver)
				values[col] = cell
			case 6:
				cell.UserEnteredValue = rwsheets.TextValue(dev.Issues)
				values[col] = cell
			case 7:
				cell.UserEnteredValue = rwsheets.TextValue(dev.Comments)
				values[col] = cell
			case 8:
				formula := fmt.Sprintf("=WORKDAY(A%d, 5, Holidays!$A$2:$A$8)", lastRow)
				cell.UserEnteredValue = rwsheets.FormulaValue(formula)
				values[col] = cell
			case 9:
				if dev.Open == "" {
					cell.UserEnteredValue = rwsheets.TextValue("Closed")
					values[col] = cell
				}
			case 12:
				if dev.Open == "" {
					cell.UserEnteredValue = rwsheets.TextValue(today)
					values[col] = cell
				}
			}
		}
		row.Values = values[:len(values)-1]
		newRows = append(newRows, row)
	}

	if len(newRows) > 0 {
		req := u.makeUpdateRequest(newRows, sheet, 0, int64(startRow))
		batchUpdate.Requests = append(batchUpdate.Requests, req)
	}

	if len(batchUpdate.Requests) > 0 {
		log.Printf("Sending %d updates to Issues sheet", len(batchUpdate.Requests))
		_, err = srv.Spreadsheets.BatchUpdate(sheet.Ssid, &batchUpdate).Do()
		if err != nil {
			log.Printf("failed to update sheet %s data - %s", sheet.Ssid, err.Error())
			return err
		}
	}

	return nil
}

func (u *UpdateManager) updateGateChecksSheet(ds map[string]*DeviationSummary) error {
	// Get our sheets service
	srv, err := u.getSheets()
	if err != nil {
		return err
	}

	u.mu.Lock()
	sheet := u.sheetData["gatechecks"]
	u.mu.Unlock()

	rows, err := rwsheets.GetSheetData(sheet.Ssid, sheet.ReadRange, srv)
	if err != nil {
		log.Printf("failed to get sheet %s data - %s", sheet.Name, err.Error())
		return err
	}

	batchUpdate := sheets.BatchUpdateSpreadsheetRequest{
		IncludeSpreadsheetInResponse: false,
	}

	lastRow := 0
	for i, row := range rows {
		values := row.Values

		auditID := values[1].FormattedValue
		if auditID == "" {
			log.Printf("Presumeably the first row (%d) in the sheet without values - exiting loop", i)
			lastRow = i
			break
		}

		vehicle := values[3].FormattedValue
		qrCode := values[4].FormattedValue
		if len(vehicle) < 5 && (len(qrCode) > 8 || qrCode == "") {
			qrCell := values[4]
			rowNum := i + 1
			qrFormula := fmt.Sprintf(`=XLOOKUP($D%d, Equipment!A:A, Equipment!B:B, "", 0)`, rowNum)
			qrCell.UserEnteredValue = rwsheets.FormulaValue(qrFormula)
			qrReq := u.makeUpdateRequest(u.makeNewCellRow(qrCell), sheet, sheet.IssuesCol-2, int64(i))
			batchUpdate.Requests = append(batchUpdate.Requests, qrReq)
		}

		dev, ok := ds[auditID]
		if !ok {
			log.Printf("Did not find audit ID %s in deviation summary", auditID)
			continue
		}
		dev.Found = true
		ds[auditID] = dev

		issues := values[6].FormattedValue
		if issues != dev.Issues {
			newCell := values[6]
			newCell.UserEnteredValue = rwsheets.TextValue(dev.Issues)
			req := u.makeUpdateRequest(u.makeNewCellRow(newCell), sheet, sheet.IssuesCol, int64(i))
			batchUpdate.Requests = append(batchUpdate.Requests, req)
		}
	}

	startRow := lastRow
	var newRows []*sheets.RowData
	for _, dev := range ds {
		if dev.Found || !strings.Contains(dev.Type, "GATE") {
			continue
		}

		if lastRow > len(rows)-1 {
			break
		}

		row := rows[lastRow]
		lastRow++
		values := row.Values

		for col, cell := range values {
			switch col {
			case 0:
				dateValue, _ := rwsheets.SerialDate(dev.Date, "1/2/2006")
				cell.UserEnteredValue = rwsheets.NumberValue(dateValue)
				values[col] = cell
			case 1:
				cell.UserEnteredValue = rwsheets.FormulaValue(dev.LinkFormula)
				values[col] = cell
			case 2:
				if dev.Division == "" {
					dev.Division = "-1"
				}
				cell.UserEnteredValue = rwsheets.TextValue(dev.Division)
				values[col] = cell
			case 3:
				cell.UserEnteredValue = rwsheets.TextValue(dev.AssetNbr)
				values[col] = cell
			case 4:
				if len(dev.AssetNbr) > 4 {
					cell.UserEnteredValue = rwsheets.TextValue(dev.AssetNbr)
					values[col] = cell
					continue
				}
				qrformula := fmt.Sprintf(`=XLOOKUP(D%d, Equipment!A:A, Equipment!B:B, "", 0)`, lastRow)
				cell.UserEnteredValue = rwsheets.FormulaValue(qrformula)
				values[col] = cell
			case 5:
				cell.UserEnteredValue = rwsheets.TextValue(dev.Driver)
				values[col] = cell
			case 6:
				cell.UserEnteredValue = rwsheets.TextValue(dev.Issues)
				values[col] = cell
			case 7:
				cell.UserEnteredValue = rwsheets.TextValue(dev.Comments)
				values[col] = cell
			case 8:
				formula := fmt.Sprintf("=WORKDAY(A%d, 5, Holidays!$A$2:$A$8)", lastRow)
				cell.UserEnteredValue = rwsheets.FormulaValue(formula)
				values[col] = cell
			}
		}
		row.Values = values[:len(values)-1]
		newRows = append(newRows, row)
	}

	if len(newRows) > 0 {
		req := u.makeUpdateRequest(newRows, sheet, 0, int64(startRow))
		batchUpdate.Requests = append(batchUpdate.Requests, req)
	}

	if len(batchUpdate.Requests) > 0 {
		log.Printf("Sending %d updates to Gate Checks sheet", len(batchUpdate.Requests))
		_, err = srv.Spreadsheets.BatchUpdate(sheet.Ssid, &batchUpdate).Do()
		if err != nil {
			log.Printf("failed to update sheet %s data - %s", sheet.Ssid, err.Error())
			return err
		}
	}

	return nil
}

func (u *UpdateManager) updateJsoSheet(ds map[string]*DeviationSummary) error {
	// Get our sheets service
	srv, err := u.getSheets()
	if err != nil {
		return err
	}

	u.mu.Lock()
	sheet := u.sheetData["jsos"]
	u.mu.Unlock()

	rows, err := rwsheets.GetSheetData(sheet.Ssid, sheet.ReadRange, srv)
	if err != nil {
		log.Printf("failed to get sheet %s data - %s", sheet.Name, err.Error())
		return err
	}

	batchUpdate := sheets.BatchUpdateSpreadsheetRequest{
		IncludeSpreadsheetInResponse: false,
	}

	lastRow := 0
	for i, row := range rows {
		values := row.Values

		auditID := values[1].FormattedValue
		if auditID == "" {
			log.Printf("Presumeably the first row (%d) in the sheet without values - exiting loop", i)
			lastRow = i
			break
		}
		dev, ok := ds[auditID]
		if !ok {
			log.Printf("Did not find audit ID %s in deviation summary", auditID)
			continue
		}
		dev.Found = true
		ds[auditID] = dev

		issues := values[5].FormattedValue
		if issues != dev.Issues && dev.Issues != "" {
			newCell := values[5]
			newCell.UserEnteredValue = rwsheets.TextValue(dev.Issues)
			req := u.makeUpdateRequest(u.makeNewCellRow(newCell), sheet, sheet.IssuesCol, int64(i))
			batchUpdate.Requests = append(batchUpdate.Requests, req)
		}
	}

	startRow := lastRow
	var newRows []*sheets.RowData
	for _, dev := range ds {
		if dev.Found || !strings.Contains(dev.Type, "JSO") {
			continue
		}

		if lastRow > len(rows)-1 {
			break
		}

		row := rows[lastRow]
		lastRow++
		values := row.Values

		for col, cell := range values {
			switch col {
			case 0:
				dateValue, _ := rwsheets.SerialDate(dev.Date, "1/2/2006")
				cell.UserEnteredValue = rwsheets.NumberValue(dateValue)
				values[col] = cell
			case 1:
				cell.UserEnteredValue = rwsheets.FormulaValue(dev.LinkFormula)
			case 2:
				if dev.Division == "" {
					dev.Division = "-1"
				}
				cell.UserEnteredValue = rwsheets.TextValue(dev.Division)
				values[col] = cell
			case 3:
				cell.UserEnteredValue = rwsheets.TextValue(dev.Driver)
				values[col] = cell
			case 4:
				cell.UserEnteredValue = rwsheets.TextValue(dev.Supervisor)
				values[col] = cell
			case 5:
				cell.UserEnteredValue = rwsheets.TextValue(dev.Issues)
				values[col] = cell
			case 6:
				cell.UserEnteredValue = rwsheets.TextValue(dev.Comments)
				values[col] = cell
			case 7:
				formula := fmt.Sprintf("=WORKDAY(A%d, 5, Holidays!$A$2:$A$8)", lastRow)
				cell.UserEnteredValue = rwsheets.FormulaValue(formula)
				values[col] = cell
			}
		}
		row.Values = values[:len(values)-1]
		newRows = append(newRows, row)
	}

	if len(newRows) > 0 {
		req := u.makeUpdateRequest(newRows, sheet, 0, int64(startRow))
		batchUpdate.Requests = append(batchUpdate.Requests, req)
	}

	if len(batchUpdate.Requests) > 0 {
		log.Printf("Sending %d updates to JSO sheet", len(batchUpdate.Requests))
		_, err = srv.Spreadsheets.BatchUpdate(sheet.Ssid, &batchUpdate).Do()
		if err != nil {
			log.Printf("failed to update sheet %s data - %s", sheet.Ssid, err.Error())
			return err
		}
	}

	return nil
}

func (u *UpdateManager) makeNewCellRow(cell *sheets.CellData) []*sheets.RowData {
	newRow := sheets.RowData{
		Values: []*sheets.CellData{cell},
	}
	newRows := []*sheets.RowData{&newRow}
	return newRows
}

func (u *UpdateManager) makeUpdateRequest(rows []*sheets.RowData, sheet Sheet, startCol, row int64) *sheets.Request {
	numCols := int64(len(rows[0].Values))

	cells := startCol + numCols

	gridRange := sheets.GridRange{
		EndColumnIndex:   cells,
		SheetId:          sheet.Gid,
		StartColumnIndex: startCol,
		StartRowIndex:    row,
		EndRowIndex:      row + int64(len(rows)),
	}

	updateCells := sheets.UpdateCellsRequest{
		Fields: "*",
		Range:  &gridRange,
		Rows:   rows,
	}

	req := sheets.Request{
		UpdateCells: &updateCells,
	}

	return &req
}
