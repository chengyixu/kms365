package auth

import (
	"context"
	"fmt"
	"net/http"
	"os"
	"time"

	"golang.org/x/oauth2"
)

// microsoftEndpoint builds the Azure AD OAuth2 endpoint for a tenant.
func microsoftEndpoint(tenantID string) oauth2.Endpoint {
	base := "https://login.microsoftonline.com/" + tenantID + "/oauth2/v2.0"
	return oauth2.Endpoint{
		AuthURL:  base + "/authorize",
		TokenURL: base + "/token",
	}
}

// GetHTTPClient returns an http.Client with OAuth2 auto-refresh if refresh_token
// + client credentials are provided, or a static-token client otherwise.
//
// Priority:
//  1. MS365_REFRESH_TOKEN + MS365_CLIENT_ID + MS365_CLIENT_SECRET + MS365_TENANT_ID → auto-refresh
//  2. MS365_ACCESS_TOKEN → static access token (backward compat)
func GetHTTPClient() (*http.Client, error) {
	refreshToken := os.Getenv("MS365_REFRESH_TOKEN")
	clientID := os.Getenv("MS365_CLIENT_ID")
	clientSecret := os.Getenv("MS365_CLIENT_SECRET")
	tenantID := os.Getenv("MS365_TENANT_ID")

	if refreshToken != "" && clientID != "" && clientSecret != "" && tenantID != "" {
		scope := os.Getenv("MS365_SCOPE")
		if scope == "" {
			scope = "offline_access User.Read"
		}
		return oauthClient(refreshToken, clientID, clientSecret, tenantID, scope), nil
	}

	// Fallback: direct access token
	token := os.Getenv("MS365_ACCESS_TOKEN")
	if token == "" {
		return nil, fmt.Errorf("MS365_REFRESH_TOKEN + MS365_CLIENT_ID + MS365_CLIENT_SECRET + MS365_TENANT_ID, or MS365_ACCESS_TOKEN environment variable is required")
	}

	return staticTokenClient(token), nil
}

func oauthClient(refreshToken, clientID, clientSecret, tenantID, scope string) *http.Client {
	cfg := oauth2.Config{
		ClientID:     clientID,
		ClientSecret: clientSecret,
		Endpoint:     microsoftEndpoint(tenantID),
		Scopes:       []string{scope},
	}

	ctx := context.WithValue(context.Background(), oauth2.HTTPClient, &http.Client{
		Timeout: 30 * time.Second,
	})

	ts := cfg.TokenSource(ctx, &oauth2.Token{RefreshToken: refreshToken})
	return oauth2.NewClient(ctx, ts)
}

func staticTokenClient(token string) *http.Client {
	ts := oauth2.StaticTokenSource(&oauth2.Token{AccessToken: token})
	return oauth2.NewClient(context.Background(), ts)
}

// GetToken returns the access token string. Kept for backward compatibility.
// Prefer GetHTTPClient() for auto-refresh.
func GetToken() (string, error) {
	token := os.Getenv("MS365_ACCESS_TOKEN")
	if token == "" {
		return "", fmt.Errorf("MS365_ACCESS_TOKEN environment variable is required")
	}
	return token, nil
}
