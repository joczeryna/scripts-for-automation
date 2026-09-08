package updates

import (
	"log"
	"safety/client"
	"sync"
	"sync/atomic"

	"github.com/pkg/errors"
	sheets "google.golang.org/api/sheets/v4"
)

type Sheet struct {
	AuditCol  int64
	Gid       int64
	Name      string
	ReadRange string
	Ssid      string
	StartRow  int64
	IssuesCol int64
	StatusCol int64
	CompCol   int64
}

type UpdateManager struct {
	sheets    atomic.Value
	sheetData map[string]Sheet

	mu sync.Mutex
}

func New() *UpdateManager {
	log.Print("Creating new updates manager now ..")

	var u UpdateManager

	u.setSheetData()

	if err := u.setupSheets(); err != nil {
		log.Fatal(err)
	}

	return &u
}

func (u *UpdateManager) setSheetData() {
	issues := Sheet{
		AuditCol:  1,
		Gid:       342524765,
		IssuesCol: 6,
		StatusCol: 9,
		CompCol:   12,
		Name:      "Issues",
		ReadRange: "Issues!A1:N",
		Ssid:      "14j7jKDtEOW8q8heUmfqsB0SLBd9r30R9pEJAj-ADtuc",
		StartRow:  1,
	}

	jsos := Sheet{
		AuditCol:  1,
		Gid:       314554508,
		IssuesCol: 5,
		StatusCol: 9,
		CompCol:   -1,
		Name:      "JSO",
		ReadRange: "JSO!A1:I",
		Ssid:      "14j7jKDtEOW8q8heUmfqsB0SLBd9r30R9pEJAj-ADtuc",
		StartRow:  1,
	}

	gc := Sheet{
		AuditCol:  1,
		Gid:       696699237,
		IssuesCol: 6,
		StatusCol: 9,
		CompCol:   12,
		Name:      "Gate Checks",
		ReadRange: "Gate Checks!A1:J",
		Ssid:      "14j7jKDtEOW8q8heUmfqsB0SLBd9r30R9pEJAj-ADtuc",
		StartRow:  1,
	}

	sheetData := map[string]Sheet{
		"issues":     issues,
		"jsos":       jsos,
		"gatechecks": gc,
	}

	u.mu.Lock()
	u.sheetData = sheetData
	u.mu.Unlock()
}

// setupSheets: Sets up the Google Sheets client.
func (u *UpdateManager) setupSheets() error {
	srv := client.Sheets()
	u.sheets.Store(srv)

	return nil
}

// getSheets: Returns the current Sheets service.
//
// Loads it from an atomic value so that it can be replaced while running without causing issues.
func (u *UpdateManager) getSheets() (*sheets.Service, error) {
	srv, ok := u.sheets.Load().(*sheets.Service)
	if !ok {
		log.Printf("u.getSheets: not a Sheets service")
		return nil, errors.New("not a Sheets service")
	}

	return srv, nil
}
