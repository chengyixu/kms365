package auth

import (
	"fmt"
	"os"
)

func GetToken() (string, error) {
	token := os.Getenv("MS365_ACCESS_TOKEN")
	if token == "" {
		return "", fmt.Errorf("MS365_ACCESS_TOKEN environment variable is required")
	}
	return token, nil
}
