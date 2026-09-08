package client

import (
	"encoding/json"
	"fmt"
	"io/ioutil"
	"log"
	"os"

	"cloud.google.com/go/firestore"
	"golang.org/x/net/context"
	"golang.org/x/oauth2"
	"golang.org/x/oauth2/google"
	"google.golang.org/api/option"
	sheets "google.golang.org/api/sheets/v4"
)

var token, credentials string

func setEnv() {
	if os.Getenv("PRODUCTION") == "false" {
		token = os.Getenv("DEV_TOKEN")
		credentials = os.Getenv("DEV_CREDENTIALS")
	} else {
		token = os.Getenv("TOKEN")
		credentials = os.Getenv("CREDENTIALS")
	}
}

// Either retrieves or creates the oauth2.Token
func getToken(config *oauth2.Config) *oauth2.Token {
	// The file token.json stores the user's access and refresh tokens, and is
	// created automatically when the authorization flow completes for the first time.
	tok, err := tokenFromFile(token)
	if err != nil {
		tok = getTokenFromWeb(config)
		saveToken(token, tok)
	}
	return tok
}

// getCredentials reads the credential variables from the env file, formatting them into the proper JSON format
// Returns that as an array of bytes to passed to Google's option.WithCredentialsJSON
func getCredentials() []byte {
	t := os.Getenv("TYPE")
	pid := os.Getenv("PROJECT_ID")
	pkid := os.Getenv("PRIVATE_KEY_ID")
	pk := os.Getenv("PRIVATE_KEY")
	ce := os.Getenv("CLIENT_EMAIL")
	cid := os.Getenv("CLIENT_ID")
	au := os.Getenv("AUTH_URI")
	tu := os.Getenv("TOKEN_URI")
	ap := os.Getenv("AUTH_PROVIDER")
	cert := os.Getenv("CLIENT_X509_CERT_URL")

	jsonCreds := fmt.Sprintf(`{
        "type": "%s",
        "project_id": "%s",
        "private_key_id": "%s",
        "private_key": "%s",
        "client_email": "%s",
        "client_id": "%s",
        "auth_uri": "%s",
        "token_uri": "%s",
        "auth_provider_x509_cert_url": "%s",
        "client_x509_cert_url": "%s"
    }`, t, pid, pkid, pk, ce, cid, au, tu, ap, cert)

	credentials := []byte(jsonCreds)
	return credentials
}

// Get the oAuth2 configuration credentials
func getConfig() *oauth2.Config {
	cred, err := ioutil.ReadFile(credentials)
	if err != nil {
		log.Fatalf("Unable to read client secret file: %v", err)
	}

	config, err := google.ConfigFromJSON(cred, "https://www.googleapis.com/auth/spreadsheets")
	if err != nil {
		log.Fatalf("Unable to parse client secret file to config: %v", err)
	}

	return config
}

// Request a token from the web, then returns the retrieved token.
func getTokenFromWeb(config *oauth2.Config) *oauth2.Token {
	authURL := config.AuthCodeURL("state-token", oauth2.AccessTypeOffline)
	fmt.Printf("Go to the following link in your browser then type the "+
		"authorization code: \n%v\n", authURL)

	var authCode string
	if _, err := fmt.Scan(&authCode); err != nil {
		log.Fatalf("Unable to read authorization code: %v", err)
	}

	tok, err := config.Exchange(context.TODO(), authCode)
	if err != nil {
		log.Fatalf("Unable to retrieve token from web: %v", err)
	}
	return tok
}

// Retrieves a token from a local file.
func tokenFromFile(file string) (*oauth2.Token, error) {
	f, err := os.Open(file)
	if err != nil {
		return nil, err
	}
	defer f.Close()
	tok := &oauth2.Token{}
	err = json.NewDecoder(f).Decode(tok)
	return tok, err
}

// Saves a token to a file path.
func saveToken(path string, token *oauth2.Token) {
	fmt.Printf("Saving credential file to: %s\n", path)
	f, err := os.OpenFile(path, os.O_RDWR|os.O_CREATE|os.O_TRUNC, 0600)
	if err != nil {
		log.Fatalf("Unable to cache oauth token: %v", err)
	}
	defer f.Close()
	json.NewEncoder(f).Encode(token)
}

// Firestore creates and returns a new firestore client
func Firestore(ctx context.Context) *firestore.Client {
	// Get our credentials
	credentials := getCredentials()

	// Set our credential variables
	pid := os.Getenv("PROJECT_ID")
	opt := option.WithCredentialsJSON(credentials)

	// Retrieve the firestore client
	client, err := firestore.NewClient(ctx, pid, opt)
	if err != nil {
		log.Fatalf("Failed to create new firestore client - %s", err.Error())
	}

	return client
}

// Sheets creates and returns a new Gmail API Service
func Sheets() *sheets.Service {
	if token == "" || credentials == "" {
		setEnv()
	}

	config := getConfig()
	tok := getToken(config)

	ctx := context.Background()
	service, err := sheets.NewService(ctx, option.WithTokenSource(config.TokenSource(ctx, tok)))
	if err != nil {
		log.Fatalf("failed to create sheets service")
		return nil
	}

	return service
}
