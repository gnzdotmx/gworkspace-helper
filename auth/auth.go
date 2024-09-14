package auth

import (
	"context"
	"encoding/json"
	"fmt"
	"io/ioutil"
	"log"
	"os"

	"golang.org/x/oauth2"
	"golang.org/x/oauth2/google"
)

// Config holds the configuration for authentication.
type Config struct {
	UseServiceAccount bool
	CredentialsFile   string
	TokenFile         string
	Scopes            []string
}

// GetClient returns an authenticated HTTP client.
func GetClient(ctx context.Context, config Config) (*oauth2.Config, *oauth2.Token, error) {
	if config.UseServiceAccount {
		return getServiceAccountClient(ctx, config)
	}
	return getOAuthClient(config)
}

// getServiceAccountClient uses a service account for authentication.
func getServiceAccountClient(ctx context.Context, config Config) (*oauth2.Config, *oauth2.Token, error) {
	data, err := ioutil.ReadFile(config.CredentialsFile)
	if err != nil {
		return nil, nil, fmt.Errorf("auth: failed to read service account file: %w", err)
	}

	creds, err := google.CredentialsFromJSON(ctx, data, config.Scopes...)
	if err != nil {
		return nil, nil, fmt.Errorf("auth: failed to parse service account credentials: %w", err)
	}

	token, err := creds.TokenSource.Token()
	if err != nil {
		return nil, nil, fmt.Errorf("auth: error when getting token: %w", err)
	}

	return nil, token, nil
}

// getOAuthClient uses OAuth2 for authentication.
func getOAuthClient(config Config) (*oauth2.Config, *oauth2.Token, error) {
	b, err := ioutil.ReadFile(config.CredentialsFile)
	if err != nil {
		return nil, nil, fmt.Errorf("auth: unable to read client secret file: %w", err)
	}

	conf, err := google.ConfigFromJSON(b, config.Scopes...)
	if err != nil {
		return nil, nil, fmt.Errorf("auth: unable to parse client secret file to config: %w", err)
	}

	tok, err := tokenFromFile(config.TokenFile)
	if err != nil {
		tok, err = getTokenFromWeb(conf)
		if err != nil {
			return nil, nil, fmt.Errorf("auth: unable to retrieve token from web: %w", err)
		}
		saveToken(config.TokenFile, tok)
	}
	return conf, tok, nil
}

func getTokenFromWeb(config *oauth2.Config) (*oauth2.Token, error) {
	authURL := config.AuthCodeURL("state-token", oauth2.AccessTypeOffline)
	fmt.Printf("auth: Go to the following link in your browser:\n%v\n", authURL)

	fmt.Print("auth: Enter the authorization code: ")
	var authCode string
	if _, err := fmt.Scan(&authCode); err != nil {
		return nil, fmt.Errorf("auth: unable to read authorization code: %w", err)
	}

	tok, err := config.Exchange(context.Background(), authCode)
	if err != nil {
		return nil, fmt.Errorf("auth: unable to retrieve token from web: %w", err)
	}
	return tok, nil
}

func tokenFromFile(file string) (*oauth2.Token, error) {
	f, err := os.Open(file)
	if err != nil {
		return nil, fmt.Errorf("auth: unable to open token file: %w", err)
	}
	defer f.Close()

	tok := &oauth2.Token{}
	err = json.NewDecoder(f).Decode(tok)
	if err != nil {
		return nil, fmt.Errorf("auth: unable to decode token file: %w", err)
	}
	return tok, nil
}

func saveToken(path string, token *oauth2.Token) {
	fmt.Printf("auth: Saving credential file to: %s\n", path)
	f, err := os.OpenFile(path, os.O_RDWR|os.O_CREATE|os.O_TRUNC, 0600)
	if err != nil {
		log.Fatalf("auth: unable to cache oauth token: %v", err)
	}
	defer f.Close()
	json.NewEncoder(f).Encode(token)
}
