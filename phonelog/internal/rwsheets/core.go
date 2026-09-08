package rwsheets

import (
	"errors"
	"time"

	sheets "google.golang.org/api/sheets/v4"
)

var (
	ErrNoData = errors.New("no sheet data found")
)

// GetSheetData: Retrieve the spreadsheet data for one sheet.
func GetSheetData(ssid, readRange string, srv *sheets.Service) ([]*sheets.RowData, error) {
	var rows []*sheets.RowData
	var ranges []string
	ranges = append(ranges, readRange)

	// Get the spreadsheet data.
	ss, err := srv.Spreadsheets.Get(ssid).Ranges(ranges...).IncludeGridData(true).Do()
	if err != nil {
		return rows, err
	}

	// Make sure we actually got at least one sheet of data.
	if len(ss.Sheets) == 0 {
		return rows, ErrNoData
	}

	// Extract the row data from the sheet.
	sheet := ss.Sheets[0]

	// Make sure we actually have data.
	if len(sheet.Data) == 0 {
		return rows, ErrNoData
	}

	grid := sheet.Data[0]
	rows = grid.RowData

	return rows, nil
}

// UpdateSheetData: Update the spreadsheet with new values.
func UpdateSheetData(ssid string, endColumnIndex, gid, startColumnIndex, startRowIndex int64, newVals []*sheets.RowData, srv *sheets.Service) error {
	var batchUpdate sheets.BatchUpdateSpreadsheetRequest
	batchUpdate.IncludeSpreadsheetInResponse = false

	gridRange := sheets.GridRange{
		EndColumnIndex:   endColumnIndex,
		SheetId:          gid,
		StartColumnIndex: startColumnIndex,
		StartRowIndex:    startRowIndex,
		EndRowIndex:      startRowIndex + int64(len(newVals)),
	}

	updateCells := sheets.UpdateCellsRequest{
		Fields: "*",
		Range:  &gridRange,
		Rows:   newVals,
	}

	request := sheets.Request{
		UpdateCells: &updateCells,
	}
	batchUpdate.Requests = append(batchUpdate.Requests, &request)
	batchUpdate.MarshalJSON()

	if _, err := srv.Spreadsheets.BatchUpdate(ssid, &batchUpdate).Do(); err != nil {
		return err
	}

	return nil
}

// RemoveRow: For removing a specific row in a Sheet.
func RemoveRow(rows []*sheets.RowData, rmvIdx int) []*sheets.RowData {
	if len(rows) == rmvIdx {
		// Let's assume we wanted to remove the last row ..
		return rows[:rmvIdx-2]
	}

	if rmvIdx+1 > len(rows) {
		return rows
	}

	return append(rows[:rmvIdx], rows[rmvIdx+1:]...)
}

// BoolValue: For updating UserEnteredValue with a boolean value.
func BoolValue(value bool) *sheets.ExtendedValue {
	return &sheets.ExtendedValue{
		BoolValue: &value,
	}
}

// FormulaValue: For updating UserEnteredValue with a formula value.
func FormulaValue(value string) *sheets.ExtendedValue {
	return &sheets.ExtendedValue{
		FormulaValue: &value,
	}
}

// TextValue: For updating UserEnteredValue with a text value.
func TextValue(value string) *sheets.ExtendedValue {
	if value == "" {
		return &sheets.ExtendedValue{}
	}
	return &sheets.ExtendedValue{
		StringValue: &value,
	}
}

// NumberValue: For updating UserEnteredValue with a number value.
// Should be used in combination with SerialDate for date values.
func NumberValue(value float64) *sheets.ExtendedValue {
	return &sheets.ExtendedValue{
		NumberValue: &value,
	}
}

// SerialDate: Returns the Google Sheets serial number for the date.
func SerialDate(value, format string) (float64, error) {
	newDate, err := time.Parse(format, value)
	if err != nil {
		return float64(0.0), err
	}

	startDate, err := time.Parse("1/2/2006", "12/30/1899")
	if err != nil {
		return float64(0.0), err
	}

	days := newDate.Sub(startDate).Hours() / 24
	return float64(days), nil
}

// TextFormat: Provides a new sheets text format.
func TextFormat(fontFamily string, fontSize int64) *sheets.TextFormat {
	return &sheets.TextFormat{
		FontFamily: fontFamily,
		FontSize:   fontSize,
	}
}

// DateFormat: Provides a sheets number format for a date value.
func DateFormat(pattern string) *sheets.NumberFormat {
	return &sheets.NumberFormat{
		Pattern: pattern,
		Type:    "DATE",
	}
}

// NumberFormat: Provides a sheets number format for a number value.
func NumberFormat(pattern string) *sheets.NumberFormat {
	return &sheets.NumberFormat{
		Pattern: pattern,
		Type:    "NUMBER",
	}
}

// CurrencyFormat: Provides the default currency formatting.
func CurrencyFormat() *sheets.NumberFormat {
	return &sheets.NumberFormat{
		Type: "CURRENCY",
	}
}

// Color: Creates a new Sheets ColorStyle with the given ARGB values.
func Color(a, b, g, r float64) *sheets.ColorStyle {
	color := sheets.Color{
		//Alpha: a,
		Blue:  b,
		Green: g,
		Red:   r,
	}
	return &sheets.ColorStyle{
		RgbColor: &color,
	}
}

// BorderConf: struct to be used to set the border style configuartion.
// BorderConf currently only supports one color for all sides that are set to true.
type BorderConf struct {
	Bottom bool
	Left   bool
	Right  bool
	Style  string
	Top    bool
	Color  *sheets.ColorStyle // Optional. Color will be set to black if not set.
}

var RIGHT_BORDER = &BorderConf{
	Right: true,
	Style: "SOLID",
}

var LEFT_BORDER = &BorderConf{
	Left:  true,
	Style: "SOLID",
}

var THICK_RIGHT_BORDER = &BorderConf{
	Right: true,
	Style: "SOLID_THICK",
}

var THICK_LEFT_BORDER = &BorderConf{
	Left:  true,
	Style: "SOLID_THICK",
}

var BLACK_COLOR = Color(1, 0, 0, 0)
var BLUE_COLOR = Color(1, 1, 0, 0)

// CellBorders: Creates a new Sheets Borders object based on the given configuration
func CellBorders(conf *BorderConf) *sheets.Borders {
	var borders sheets.Borders
	style := conf.Style
	if style == "" {
		style = "SOLID"
	}

	color := conf.Color
	if color == nil {
		color = BLACK_COLOR
	}
	if conf.Bottom {
		borders.Bottom = &sheets.Border{
			ColorStyle: color,
			Style:      style,
		}
	}
	if conf.Left {
		borders.Left = &sheets.Border{
			ColorStyle: color,
			Style:      style,
		}
	}
	if conf.Right {
		borders.Right = &sheets.Border{
			ColorStyle: color,
			Style:      style,
		}
	}
	if conf.Top {
		borders.Top = &sheets.Border{
			ColorStyle: color,
			Style:      style,
		}
	}
	return &borders
}

// Styler is to be used to create new cells with styling.
type Styler struct {
	fontBold            bool
	fontColor           *sheets.ColorStyle
	fontFamily          string
	fontSize            int64
	datePattern         string
	numberPattern       string
	horizontalAlignment string
	verticalAlignment   string
}

// NewStyler: Returns a new styler with developer preferred settings.
// To override any of style settings, please use the applicable function for the styler.
func NewStyler() *Styler {
	return &Styler{
		fontBold:            false,
		fontColor:           BLACK_COLOR,
		fontFamily:          "Verdana",
		fontSize:            int64(10),
		datePattern:         "M/d/yyyy",
		numberPattern:       "#,##0.00_);-#,##0.00",
		horizontalAlignment: "LEFT",
		verticalAlignment:   "MIDDLE",
	}
}

// Sets whether the styler should make the font bold.
func (s *Styler) FontColor(c *sheets.ColorStyle) *Styler {
	s.fontColor = c
	return s
}

// Sets whether the styler should make the font bold.
func (s *Styler) FontBold(bold bool) *Styler {
	s.fontBold = bold
	return s
}

// Sets the stylers default font family.
func (s *Styler) FontFamily(fontFamily string) *Styler {
	s.fontFamily = fontFamily
	return s
}

// Sets the stylers default font size.
func (s *Styler) FontSize(fontSize int64) *Styler {
	s.fontSize = fontSize
	return s
}

// Sets the stylers date pattern to use when parsing dates.
func (s *Styler) DatePattern(pattern string) *Styler {
	s.datePattern = pattern
	return s
}

// Sets the stylers number pattern to use when creating number value cells.
func (s *Styler) NumberPattern(pattern string) *Styler {
	s.numberPattern = pattern
	return s
}

// Sets the stylers horizontal alignment to use when creating text formats.
func (s *Styler) HorizontalAlignment(alignment string) *Styler {
	if alignment == "" {
		alignment = "LEFT"
	}
	if alignment == "MIDDLE" {
		alignment = "CENTER"
	}
	s.horizontalAlignment = alignment
	return s
}

// Sets the stylers vertical alignment to use when creating text formats.
func (s *Styler) VerticalAlignment(alignment string) *Styler {
	if alignment == "" {
		alignment = "MIDDLE"
	}
	s.verticalAlignment = alignment
	return s
}

// TextFormat: Provides a new sheets Text Format using the stylers settings.
func (s *Styler) TextFormat() *sheets.TextFormat {
	return &sheets.TextFormat{
		Bold:                 s.fontBold,
		FontFamily:           s.fontFamily,
		FontSize:             s.fontSize,
		ForegroundColorStyle: s.fontColor,
	}
}

// DateFormat: Provides a sheets number format for a date value using the stylers settings.
func (s *Styler) DateFormat() *sheets.NumberFormat {
	return &sheets.NumberFormat{
		Pattern: s.datePattern,
		Type:    "DATE",
	}
}

// NumberFormat: Provides a sheets number format for a number value using the stylers settings.
func (s *Styler) NumberFormat() *sheets.NumberFormat {
	return &sheets.NumberFormat{
		Pattern: s.numberPattern,
		Type:    "NUMBER",
	}
}

// CurrencyFormat: Provides the default currency formatting using the styler.
func (s *Styler) CurrencyFormat() *sheets.NumberFormat {
	return &sheets.NumberFormat{
		Type: "CURRENCY",
	}
}

// AccountingFormat: Provides the default accounting formatting using the styler.
func (s *Styler) AccountingFormat() *sheets.NumberFormat {
	return &sheets.NumberFormat{
		Pattern: `_("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)`,
		Type:    "NUMBER",
	}
}

// TextCell: Creates a new sheets text cell using the stylers settings for the formatting.
func (s *Styler) TextCell(value string, borders *BorderConf) *sheets.CellData {
	format := sheets.CellFormat{
		HorizontalAlignment: s.horizontalAlignment,
		TextFormat:          s.TextFormat(),
		VerticalAlignment:   s.verticalAlignment,
	}
	if borders != nil {
		format.Borders = CellBorders(borders)
	}

	return &sheets.CellData{
		UserEnteredFormat: &format,
		UserEnteredValue:  TextValue(value),
	}
}

// BoolCell: Creates a new sheets bool cell using the stylers settings for the formatting.
func (s *Styler) BoolCell(value bool, borders *BorderConf) *sheets.CellData {
	format := sheets.CellFormat{
		HorizontalAlignment: s.horizontalAlignment,
		TextFormat:          s.TextFormat(),
		VerticalAlignment:   s.verticalAlignment,
	}
	if borders != nil {
		format.Borders = CellBorders(borders)
	}

	return &sheets.CellData{
		UserEnteredFormat: &format,
		UserEnteredValue:  BoolValue(value),
	}
}

// CheckBoxCell: Creates a new sheets checkbox cell using the stylers settings for the formatting.
func (s *Styler) CheckBoxCell(value bool, borders *BorderConf) *sheets.CellData {
	bc := sheets.BooleanCondition{
		Type: "BOOLEAN",
	}
	dv := sheets.DataValidationRule{
		Condition: &bc,
		Strict:    true,
	}
	format := sheets.CellFormat{
		HorizontalAlignment: s.horizontalAlignment,
		TextFormat:          s.TextFormat(),
		VerticalAlignment:   s.verticalAlignment,
	}
	if borders != nil {
		format.Borders = CellBorders(borders)
	}

	return &sheets.CellData{
		DataValidation:    &dv,
		UserEnteredFormat: &format,
		UserEnteredValue:  BoolValue(value),
	}
}

// NumberCell: Creates a new sheets text cell using the stylers settings for the formatting.
func (s *Styler) NumberCell(value float64, borders *BorderConf) *sheets.CellData {
	format := sheets.CellFormat{
		HorizontalAlignment: s.horizontalAlignment,
		NumberFormat:        s.NumberFormat(),
		TextFormat:          s.TextFormat(),
		VerticalAlignment:   s.verticalAlignment,
	}
	if borders != nil {
		format.Borders = CellBorders(borders)
	}

	return &sheets.CellData{
		UserEnteredFormat: &format,
		UserEnteredValue:  NumberValue(value),
	}
}

// AccountingCell: Creates a new sheets accounting cell using the stylers settings for the formatting.
func (s *Styler) AccountingCell(value float64, borders *BorderConf) *sheets.CellData {
	format := sheets.CellFormat{
		HorizontalAlignment: s.horizontalAlignment,
		NumberFormat:        s.AccountingFormat(),
		TextFormat:          s.TextFormat(),
		VerticalAlignment:   s.verticalAlignment,
	}
	if borders != nil {
		format.Borders = CellBorders(borders)
	}

	return &sheets.CellData{
		UserEnteredFormat: &format,
		UserEnteredValue:  NumberValue(value),
	}
}

// DateCell: Creates a new sheets date cell using the stylers settings for the formatting.
func (s *Styler) DateCell(date, layout string, borders *BorderConf) *sheets.CellData {
	serialDate, err := SerialDate(date, layout)
	if err != nil {
		return s.TextCell(date, borders)
	}

	format := sheets.CellFormat{
		HorizontalAlignment: s.horizontalAlignment,
		NumberFormat:        s.DateFormat(),
		TextFormat:          s.TextFormat(),
		VerticalAlignment:   s.verticalAlignment,
	}

	if borders != nil {
		format.Borders = CellBorders(borders)
	}

	return &sheets.CellData{
		UserEnteredFormat: &format,
		UserEnteredValue:  NumberValue(serialDate),
	}
}

// CreateHeaderRow: Creates the header row with the given header values.
func (s *Styler) CreateHeaderRow(headerValues []string, borders *BorderConf) []*sheets.RowData {
	var rows []*sheets.RowData
	var headerCells []*sheets.CellData

	for _, val := range headerValues {
		headerCells = append(headerCells, s.TextCell(val, borders))
	}

	headerRow := sheets.RowData{
		Values: headerCells,
	}
	rows = append(rows, &headerRow)

	return rows
}
