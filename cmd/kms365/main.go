package main

import (
	"fmt"
	"os"

	"github.com/minervacap2022/klik-ms365-cli/internal/api"
	"github.com/minervacap2022/klik-ms365-cli/internal/auth"
	"github.com/minervacap2022/klik-ms365-cli/internal/output"
	"github.com/spf13/cobra"
)

var rootCmd = &cobra.Command{
	Use:   "kms365",
	Short: "Microsoft 365 CLI for KLIK platform",
	Long:  "Command-line interface for Microsoft Graph API. Auth via MS365_ACCESS_TOKEN env var.",
}

func getClient() *api.Client {
	token, err := auth.GetToken()
	if err != nil {
		output.Error(err.Error())
		os.Exit(1)
	}
	return api.NewClient(token)
}

// --- mail commands ---

var mailCmd = &cobra.Command{
	Use:   "mail",
	Short: "Manage emails",
}

var mailListCmd = &cobra.Command{
	Use:   "list",
	Short: "List emails",
	RunE: func(cmd *cobra.Command, args []string) error {
		client := getClient()
		limit, _ := cmd.Flags().GetInt("limit")
		folder, _ := cmd.Flags().GetString("folder")

		path := fmt.Sprintf("/me/mailFolders/%s/messages?$top=%d&$select=subject,from,receivedDateTime,isRead,bodyPreview", folder, limit)
		result, err := client.Get(path)
		if err != nil {
			return err
		}
		output.RawJSON(result)
		return nil
	},
}

var mailReadCmd = &cobra.Command{
	Use:   "read",
	Short: "Read an email",
	RunE: func(cmd *cobra.Command, args []string) error {
		client := getClient()
		id, _ := cmd.Flags().GetString("id")

		result, err := client.Get(fmt.Sprintf("/me/messages/%s", id))
		if err != nil {
			return err
		}
		output.RawJSON(result)
		return nil
	},
}

var mailSendCmd = &cobra.Command{
	Use:   "send",
	Short: "Send an email",
	RunE: func(cmd *cobra.Command, args []string) error {
		client := getClient()
		to, _ := cmd.Flags().GetString("to")
		subject, _ := cmd.Flags().GetString("subject")
		body, _ := cmd.Flags().GetString("body")

		payload := map[string]interface{}{
			"message": map[string]interface{}{
				"subject": subject,
				"body": map[string]string{
					"contentType": "Text",
					"content":     body,
				},
				"toRecipients": []map[string]interface{}{
					{
						"emailAddress": map[string]string{
							"address": to,
						},
					},
				},
			},
			"saveToSentItems": true,
		}

		_, err := client.Post("/me/sendMail", payload)
		if err != nil {
			return err
		}
		output.JSON(map[string]interface{}{"ok": true, "message": "Email sent"})
		return nil
	},
}

var mailReplyCmd = &cobra.Command{
	Use:   "reply",
	Short: "Reply to an email",
	RunE: func(cmd *cobra.Command, args []string) error {
		client := getClient()
		id, _ := cmd.Flags().GetString("id")
		body, _ := cmd.Flags().GetString("body")

		payload := map[string]interface{}{
			"comment": body,
		}

		_, err := client.Post(fmt.Sprintf("/me/messages/%s/reply", id), payload)
		if err != nil {
			return err
		}
		output.JSON(map[string]interface{}{"ok": true, "message": "Reply sent"})
		return nil
	},
}

// --- calendar commands ---

var calendarCmd = &cobra.Command{
	Use:   "calendar",
	Short: "Manage calendar events",
}

var calendarListCmd = &cobra.Command{
	Use:   "list",
	Short: "List calendar events",
	RunE: func(cmd *cobra.Command, args []string) error {
		client := getClient()
		limit, _ := cmd.Flags().GetInt("limit")

		path := fmt.Sprintf("/me/events?$top=%d&$select=subject,start,end,location,organizer,isAllDay&$orderby=start/dateTime", limit)
		result, err := client.Get(path)
		if err != nil {
			return err
		}
		output.RawJSON(result)
		return nil
	},
}

var calendarCreateCmd = &cobra.Command{
	Use:   "create",
	Short: "Create a calendar event",
	RunE: func(cmd *cobra.Command, args []string) error {
		client := getClient()
		subject, _ := cmd.Flags().GetString("subject")
		start, _ := cmd.Flags().GetString("start")
		end, _ := cmd.Flags().GetString("end")
		location, _ := cmd.Flags().GetString("location")

		payload := map[string]interface{}{
			"subject": subject,
			"start": map[string]string{
				"dateTime": start,
				"timeZone": "UTC",
			},
			"end": map[string]string{
				"dateTime": end,
				"timeZone": "UTC",
			},
		}
		if location != "" {
			payload["location"] = map[string]string{"displayName": location}
		}

		result, err := client.Post("/me/events", payload)
		if err != nil {
			return err
		}
		output.RawJSON(result)
		return nil
	},
}

var calendarDeleteCmd = &cobra.Command{
	Use:   "delete",
	Short: "Delete a calendar event",
	RunE: func(cmd *cobra.Command, args []string) error {
		client := getClient()
		id, _ := cmd.Flags().GetString("id")

		err := client.Delete(fmt.Sprintf("/me/events/%s", id))
		if err != nil {
			return err
		}
		output.JSON(map[string]interface{}{"ok": true, "message": "Event deleted"})
		return nil
	},
}

// --- event commands ---

var eventCmd = &cobra.Command{
	Use:   "event",
	Short: "Manage event responses",
}

var eventRespondCmd = &cobra.Command{
	Use:   "respond",
	Short: "Respond to an event (accept/decline/tentative)",
	RunE: func(cmd *cobra.Command, args []string) error {
		client := getClient()
		id, _ := cmd.Flags().GetString("id")
		response, _ := cmd.Flags().GetString("response")
		comment, _ := cmd.Flags().GetString("comment")

		payload := map[string]interface{}{
			"sendResponse": true,
		}
		if comment != "" {
			payload["comment"] = comment
		}

		_, err := client.Post(fmt.Sprintf("/me/events/%s/%s", id, response), payload)
		if err != nil {
			return err
		}
		output.JSON(map[string]interface{}{"ok": true, "message": fmt.Sprintf("Event %sed", response)})
		return nil
	},
}

// --- teams commands ---

var teamsCmd = &cobra.Command{
	Use:   "teams",
	Short: "Manage Microsoft Teams",
}

var teamsMessageCmd = &cobra.Command{
	Use:   "message",
	Short: "Teams messaging",
}

var teamsMessageSendCmd = &cobra.Command{
	Use:   "send",
	Short: "Send a message to a Teams channel",
	RunE: func(cmd *cobra.Command, args []string) error {
		client := getClient()
		team, _ := cmd.Flags().GetString("team")
		channel, _ := cmd.Flags().GetString("channel")
		text, _ := cmd.Flags().GetString("text")

		payload := map[string]interface{}{
			"body": map[string]string{
				"content": text,
			},
		}

		result, err := client.Post(fmt.Sprintf("/teams/%s/channels/%s/messages", team, channel), payload)
		if err != nil {
			return err
		}
		output.RawJSON(result)
		return nil
	},
}

// --- todo commands ---

var todoCmd = &cobra.Command{
	Use:   "todo",
	Short: "Manage Microsoft To Do",
}

var todoListCmd = &cobra.Command{
	Use:   "list",
	Short: "List todo tasks",
	RunE: func(cmd *cobra.Command, args []string) error {
		client := getClient()
		listName, _ := cmd.Flags().GetString("list")

		if listName != "" {
			result, err := client.Get(fmt.Sprintf("/me/todo/lists?$filter=displayName eq '%s'", listName))
			if err != nil {
				return err
			}
			output.RawJSON(result)
		} else {
			result, err := client.Get("/me/todo/lists")
			if err != nil {
				return err
			}
			output.RawJSON(result)
		}
		return nil
	},
}

var todoCreateCmd = &cobra.Command{
	Use:   "create",
	Short: "Create a todo task",
	RunE: func(cmd *cobra.Command, args []string) error {
		client := getClient()
		listID, _ := cmd.Flags().GetString("list-id")
		title, _ := cmd.Flags().GetString("title")

		payload := map[string]interface{}{
			"title": title,
		}

		result, err := client.Post(fmt.Sprintf("/me/todo/lists/%s/tasks", listID), payload)
		if err != nil {
			return err
		}
		output.RawJSON(result)
		return nil
	},
}

// --- onedrive commands ---

var onedriveCmd = &cobra.Command{
	Use:   "onedrive",
	Short: "Manage OneDrive files",
}

var onedriveListCmd = &cobra.Command{
	Use:   "list",
	Short: "List OneDrive files",
	RunE: func(cmd *cobra.Command, args []string) error {
		client := getClient()
		path, _ := cmd.Flags().GetString("path")

		var apiPath string
		if path != "" && path != "/" {
			apiPath = fmt.Sprintf("/me/drive/root:/%s:/children", path)
		} else {
			apiPath = "/me/drive/root/children"
		}

		result, err := client.Get(apiPath)
		if err != nil {
			return err
		}
		output.RawJSON(result)
		return nil
	},
}

func init() {
	// mail
	mailListCmd.Flags().Int("limit", 10, "Max emails")
	mailListCmd.Flags().String("folder", "inbox", "Mail folder")

	mailReadCmd.Flags().String("id", "", "Message ID")
	mailReadCmd.MarkFlagRequired("id")

	mailSendCmd.Flags().String("to", "", "Recipient email")
	mailSendCmd.Flags().String("subject", "", "Subject")
	mailSendCmd.Flags().String("body", "", "Body text")
	mailSendCmd.MarkFlagRequired("to")
	mailSendCmd.MarkFlagRequired("subject")
	mailSendCmd.MarkFlagRequired("body")

	mailReplyCmd.Flags().String("id", "", "Message ID to reply to")
	mailReplyCmd.Flags().String("body", "", "Reply text")
	mailReplyCmd.MarkFlagRequired("id")
	mailReplyCmd.MarkFlagRequired("body")

	mailCmd.AddCommand(mailListCmd, mailReadCmd, mailSendCmd, mailReplyCmd)

	// calendar
	calendarListCmd.Flags().Int("limit", 10, "Max events")

	calendarCreateCmd.Flags().String("subject", "", "Event subject")
	calendarCreateCmd.Flags().String("start", "", "Start datetime (ISO 8601)")
	calendarCreateCmd.Flags().String("end", "", "End datetime (ISO 8601)")
	calendarCreateCmd.Flags().String("location", "", "Location")
	calendarCreateCmd.MarkFlagRequired("subject")
	calendarCreateCmd.MarkFlagRequired("start")
	calendarCreateCmd.MarkFlagRequired("end")

	calendarDeleteCmd.Flags().String("id", "", "Event ID")
	calendarDeleteCmd.MarkFlagRequired("id")

	calendarCmd.AddCommand(calendarListCmd, calendarCreateCmd, calendarDeleteCmd)

	// event
	eventRespondCmd.Flags().String("id", "", "Event ID")
	eventRespondCmd.Flags().String("response", "", "Response: accept, decline, tentativelyAccept")
	eventRespondCmd.Flags().String("comment", "", "Optional comment")
	eventRespondCmd.MarkFlagRequired("id")
	eventRespondCmd.MarkFlagRequired("response")
	eventCmd.AddCommand(eventRespondCmd)

	// teams
	teamsMessageSendCmd.Flags().String("team", "", "Team ID")
	teamsMessageSendCmd.Flags().String("channel", "", "Channel ID")
	teamsMessageSendCmd.Flags().String("text", "", "Message text")
	teamsMessageSendCmd.MarkFlagRequired("team")
	teamsMessageSendCmd.MarkFlagRequired("channel")
	teamsMessageSendCmd.MarkFlagRequired("text")
	teamsMessageCmd.AddCommand(teamsMessageSendCmd)
	teamsCmd.AddCommand(teamsMessageCmd)

	// todo
	todoListCmd.Flags().String("list", "", "List name filter")
	todoCreateCmd.Flags().String("list-id", "", "Todo list ID")
	todoCreateCmd.Flags().String("title", "", "Task title")
	todoCreateCmd.MarkFlagRequired("list-id")
	todoCreateCmd.MarkFlagRequired("title")
	todoCmd.AddCommand(todoListCmd, todoCreateCmd)

	// onedrive
	onedriveListCmd.Flags().String("path", "", "Folder path")
	onedriveCmd.AddCommand(onedriveListCmd)

	rootCmd.AddCommand(mailCmd, calendarCmd, eventCmd, teamsCmd, todoCmd, onedriveCmd)
}

func main() {
	if err := rootCmd.Execute(); err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
}
