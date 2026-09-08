package updates

import (
	"fmt"
	"log"
	"os"
	"strings"

	excelize "github.com/xuri/excelize/v2"
)

type Deviation struct {
	Subsidiary string
	Location   string
	Worker     string
	WorkerNbr  string
	AssetNbr   string
	Supervisor string
	Auditor    string
	AuditDate  string
	AuditId    string
	AuditLink  string
	AuditName  string
	Question   string
	DueDate    string
	DaysTilDue string
	Status     string
}

type DeviationSummary struct {
	Date        string
	AuditId     string
	Location    string
	Division    string
	AssetNbr    string
	Driver      string
	Supervisor  string
	Issues      string
	Comments    string
	DateDue     string
	LinkFormula string
	Closed      string
	Open        string
	Type        string
	Found       bool
}

func (u *UpdateManager) UpdateReport() error {
	ds := u.getSummary()
	log.Printf("Summarized deviations: %d", len(ds))

	log.Printf("Updating Issues sheet now")
	if err := u.updateIssuesSheet(ds); err != nil {
		return err
	}

	log.Printf("Updating Gate Checks sheet now")
	if err := u.updateGateChecksSheet(ds); err != nil {
		return err
	}

	log.Printf("Updating JSO sheet now")
	if err := u.updateJsoSheet(ds); err != nil {
		return err
	}
	return nil
}

func (u *UpdateManager) getSummary() map[string]*DeviationSummary {
	devs := u.readFile()
	log.Printf("Deviations found: %d", len(devs))
	ds := u.makeDeviationSummary(devs)
	return ds
}

func (u *UpdateManager) readFile() []*Deviation {
	var deviations []*Deviation

	path := os.Getenv("FILE_PATH")
	f, err := excelize.OpenFile(path)
	if err != nil {
		log.Print(err)
		return deviations
	}

	defer func() {
		// Close the spreadsheet.
		if err := f.Close(); err != nil {
			log.Print(err)
		}
	}()

	sheetname := "Export"

	// Get all the rows in the sheet.
	rows, err := f.GetRows(sheetname)
	if err != nil {
		log.Print(err)
		return deviations

	}

	for i, row := range rows {
		if i < 1 {
			continue
		}

		var d Deviation
		for j, cell := range row {
			switch j {
			case 0:
				d.Subsidiary = cell
			case 1:
				d.Location = cell
			case 2:
				d.Worker = cell
			case 3:
				d.WorkerNbr = cell
				if d.WorkerNbr == "-1" {
					d.WorkerNbr = ""
				}
			case 4:
				d.AssetNbr = cell
			case 5:
				d.Supervisor = cell
			case 6:
				d.Auditor = cell
			case 7:
				dateArr := strings.Split(cell, " ")
				if len(dateArr) > 0 {
					d.AuditDate = dateArr[0]
				}
			case 8:
				d.AuditId = cell
			case 9:
				d.AuditLink = cell
			case 10:
				d.AuditName = cell
			case 11:
				question, ok := questions[cell]
				if ok {
					d.Question = question
				} else {
					d.Question = cell
				}
			case 12:
				d.DueDate = cell
			case 13:
				d.DaysTilDue = cell
			case 14:
				d.Status = cell
			}
		}
		if d.AuditId == "" || d.AuditName == "FORKLIFT DAILY INSPECTION" {
			continue
		}
		deviations = append(deviations, &d)
	}

	return deviations
}

func (u *UpdateManager) makeDeviationSummary(devs []*Deviation) map[string]*DeviationSummary {
	devSummary := make(map[string]*DeviationSummary)

	for _, d := range devs {
		ds, ok := devSummary[d.AuditId]
		if ok {
			u.combineData(ds, d)
			log.Printf("%v", *ds)
			devSummary[d.AuditId] = ds
			continue
		}
		dev := DeviationSummary{
			Date:        d.AuditDate,
			AuditId:     d.AuditId,
			Location:    d.Location,
			Division:    locs[d.Location],
			AssetNbr:    d.AssetNbr,
			Driver:      d.Worker,
			Supervisor:  d.Supervisor,
			LinkFormula: fmt.Sprintf(`=HYPERLINK("%s", "%s")`, d.AuditLink, d.AuditId),
			Type:        d.AuditName,
		}
		if d.AuditName == "JSO CONSTRUCTION" && dev.Division == "" {
			dev.Division = auditors[d.Auditor]
		}
		if d.Status == "ACTIVE" {
			dev.Open = d.Question
			dev.Issues = d.Question
		} else {
			dev.Closed = d.Question
			dev.Issues = "RESOLVED:\n" + d.Question
		}
		//log.Printf("%v", dev)
		devSummary[d.AuditId] = &dev
	}

	return devSummary
}

func (u *UpdateManager) combineData(ds *DeviationSummary, d *Deviation) {
	if ds.AssetNbr == "" {
		ds.AssetNbr = d.AssetNbr
	}
	if ds.Date == "" {
		ds.Date = d.AuditDate
	}
	if ds.Location == "" {
		ds.Location = d.Location
		ds.Division = locs[d.Location]
	}
	if ds.Division == "" && d.AuditName == "JSO CONSTRUCTION" {
		ds.Division = auditors[d.Auditor]
	}

	ds.Driver = addValue(d.Worker, ds.Driver)
	ds.Supervisor = addValue(d.Supervisor, ds.Supervisor)

	if d.Status == "ACTIVE" {
		ds.Open = addValue(d.Question, ds.Open)
	} else {
		ds.Closed = addValue(d.Question, ds.Closed)
	}

	ds.Issues = ds.Open
	if ds.Closed != "" {
		if ds.Issues == "" {
			ds.Issues = "RESOLVED:\n" + ds.Closed
		} else {
			ds.Issues = ds.Issues + "\nRESOLVED:\n" + ds.Closed
		}
	}
}

func addValue(newVal, currVal string) string {
	if currVal == "" {
		return newVal
	}
	if newVal == "" {
		return currVal
	}
	if !strings.Contains(currVal, newVal) {
		currVal = currVal + "\n" + newVal
	}
	return currVal
}
