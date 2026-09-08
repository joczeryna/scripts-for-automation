package rwsheets

import (
	"encoding/json"
	"errors"
	"fmt"
	"log"
	"os"

	"golang.org/x/net/context"
	"golang.org/x/oauth2"
	"golang.org/x/oauth2/google"
	"google.golang.org/api/option"
	sheets "google.golang.org/api/sheets/v4"
)

var (
	ErrReadCredentials = errors.New("unable to read contents of credential file")
	ErrReadToken       = errors.New("unable to read contents of token file")
	ErrConfig          = errors.New("failed to create oauth2.Config")
)

// NewSheetsService: Creates a new Google Sheets Service.
//
// credentialFile should be the file path to your GCP oAuth2 client credential key.
// tokenFile should the file path to the file containing the key created with the given scopes
// scope should be the oAuth scope needed for the Sheets service
//
// Sheets scope options:
// https://www.googleapis.com/auth/drive - See, edit, create, and delete all of your Google Drive files
// https://www.googleapis.com/auth/drive.file - See, edit, create, and delete only the specific Google Drive files you use with this app
// https://www.googleapis.com/auth/drive.readonly - See and download all your Google Drive files
// https://www.googleapis.com/auth/spreadsheets - See, edit, create, and delete all your Google Sheets spreadsheets
// https://www.googleapis.com/auth/spreadsheets.readonly - See all your Google Sheets spreadsheets
func NewSheetsService(ctx context.Context, credentialFile, tokenFile string, scope ...string) (*sheets.Service, error) {
	config, err := getConfig(credentialFile, scope...)
	if err != nil {
		return nil, err
	}

	token, err := getToken(config, tokenFile)
	if err != nil {
		return nil, err
	}

	return sheets.NewService(ctx, option.WithTokenSource(config.TokenSource(ctx, token)))
}

// getConfig: Retrieves the oauth2.Config using the given client credientals and scope.
func getConfig(credentialFile string, scope ...string) (*oauth2.Config, error) {
	// Read the users client credential file.
	cred, err := os.ReadFile(credentialFile)
	if err != nil {
		log.Printf("getConfig: unable to read client secret file: %v", err)
		return nil, ErrReadCredentials
	}

	// If modifying these scopes, delete your previously saved token.json.
	config, err := google.ConfigFromJSON(cred, scope...)
	if err != nil {
		log.Printf("getConfig: failed to create oauth2.Config: %v", err)
		return nil, ErrConfig
	}

	return config, nil
}

// getToken: Either retrieves or creates the oauth2.Token.
func getToken(config *oauth2.Config, token string) (*oauth2.Token, error) {
	// The given token file stores the user's access and refresh tokens, and is
	// created automatically when the authorization flow completes for the first time.
	tok, err := tokenFromFile(token)
	if err != nil {
		tok = getTokenFromWeb(config)
		saveToken(token, tok)
	}
	return tok, nil
}

// tokenFromFile: Retrieves a token from a local file.
func tokenFromFile(file string) (*oauth2.Token, error) {
	// Open the token file.
	f, err := os.Open(file)
	if err != nil {
		// Couldn't open the file.
		return nil, ErrReadToken
	}
	defer f.Close()

	// Create a new oauth2.Token.
	tok := &oauth2.Token{}

	// Decode the contents of the file into the tok object.
	err = json.NewDecoder(f).Decode(tok)
	return tok, err
}

// getTokenFromWeb: Request a token from the web, then returns the retrieved token.
func getTokenFromWeb(config *oauth2.Config) *oauth2.Token {
	// Create a new authorization URL.
	authURL := config.AuthCodeURL("state-token", oauth2.AccessTypeOffline)
	fmt.Printf("Go to the following link in your browser then type the "+
		"authorization code: \n%v\n", authURL)

	var authCode string
	// Read the auth code from the terminal
	if _, err := fmt.Scan(&authCode); err != nil {
		log.Fatalf("getTokenFromWeb: unable to read authorization code: %v", err)
	}

	// Create the oauth2.Token.
	tok, err := config.Exchange(context.TODO(), authCode)
	if err != nil {
		log.Fatalf("getTokenFromWeb: unable to retrieve token from web: %v", err)
	}
	return tok
}

// saveToken: Saves a token to a file path.
func saveToken(path string, token *oauth2.Token) {
	log.Printf("saveToken: saving file to: %s\n", path)

	// Open or create the token file at the given path.
	f, err := os.OpenFile(path, os.O_RDWR|os.O_CREATE|os.O_TRUNC, 0600)
	if err != nil {
		log.Printf("saveToken: unable to cache oauth token at path %s: %v", path, err)
	}
	defer f.Close()

	// Add the contents of the token to the file.
	json.NewEncoder(f).Encode(token)
}
