package main

import (
	"context"
	"fmt"
	"log"
	"time"

	"github.com/gnzdotmx/gworkspace-helper/auth"
	"github.com/gnzdotmx/gworkspace-helper/gDriveHelper"
	"github.com/gnzdotmx/gworkspace-helper/gMeetHelper"
	"github.com/gnzdotmx/gworkspace-helper/gdocsHelper"
	"google.golang.org/api/docs/v1"
)

func main() {
	ctx := context.Background()

	// Configure authentication
	authConfig := auth.Config{
		UseServiceAccount: false,                   // Set to true to use a service account
		CredentialsFile:   "auth/credentials.json", // Path to your credentials file
		TokenFile:         "auth/token.json",       // Path to your token file
		Scopes: []string{
			"https://www.googleapis.com/auth/drive",
			"https://www.googleapis.com/auth/documents",
			"https://www.googleapis.com/auth/calendar",
		},
	}

	// ************** CREATE A GOOGLE DOC
	doc, err := gdocsHelper.CreateGoogleDoc(ctx, authConfig, "Sample Document")
	if err != nil {
		log.Fatalf("main: unable to create document: %v", err)
	}
	fmt.Printf("Created document with ID: %s\n", doc.DocumentId)

	// ************** ADD TEXT TO A DOCUMENT
	err = gdocsHelper.AddText(ctx, authConfig, doc.DocumentId, "Hello, World!")
	if err != nil {
		log.Fatalf("main: unable to add text to document: %v", err)
	}
	fmt.Println("Added text to the document.")

	// ************** CREATE A FOLDER IN GOOGLE DRIVE
	folder, err := gDriveHelper.CreateFolder(ctx, authConfig, "Sample Folder")
	if err != nil {
		log.Fatalf("main: unable to create folder: %v", err)
	}
	fmt.Printf("Created folder with ID: %s\n", folder.Id)

	// ************** COPY THE GOOGLE DOC TO A FOLDER
	// Copy the Google Doc to the folder
	copiedFile, err := gDriveHelper.CopyFileToFolder(ctx, authConfig, doc.DocumentId, folder.Id)
	if err != nil {
		log.Fatalf("main: unable to copy document to folder: %v", err)
	}
	fmt.Printf("Copied document to folder with ID: %s\n", copiedFile.Id)

	// ************** CREATE A GOOGLE CALENDAR EVENT
	startTime := time.Now()
	endTime := startTime.Add(1 * time.Hour)
	event, err := gMeetHelper.CreateCalendarEvent(ctx, authConfig, "Meeting", "Virtual", "Discuss project updates", startTime, endTime)
	if err != nil {
		log.Fatalf("main: unable to create calendar event: %v", err)
	}
	fmt.Printf("Created event with ID: %s\n", event.Id)

	// ************** ADD ATTENDEES TO THE EVENT
	attendees := []string{"user@gmail.com"}
	err = gMeetHelper.AddAttendeesToEvent(ctx, authConfig, event.Id, attendees)
	if err != nil {
		log.Fatalf("main: unable to add attendees to event: %v", err)
	}
	fmt.Println("Added attendees to the event.")

	// Attach the copied document to the event
	err = gMeetHelper.AttachFileToEvent(ctx, authConfig, event.Id, copiedFile.Id)
	if err != nil {
		log.Fatalf("main: unable to attach file to event: %v", err)
	}
	fmt.Println("Attached file to the event.")

	// ************** COPY A GOOGLE DOC
	copiedFile, err = gdocsHelper.MakeCopyOfGoogleDoc(ctx, authConfig, doc.DocumentId, "Copy of Document")
	if err != nil {
		log.Fatalf("Unable to copy document: %v", err)
	}
	fmt.Printf("Copied document ID: %s\n", copiedFile.Id)

	// EXPORT A GOOGLE DOC AS TEXT
	content, err := gdocsHelper.ExportGoogleDocAsText(ctx, authConfig, doc.DocumentId)
	if err != nil {
		log.Fatalf("Unable to export document: %v", err)
	}
	fmt.Printf("Document content:\n%s\n", content)

	// ************** ADD A TABLE TO A GDOC FILE
	err = gdocsHelper.AddTable(ctx, authConfig, copiedFile.Id, 3, 3)
	if err != nil {
		log.Fatalf("Unable to add table: %v", err)
	}
	fmt.Println("Added table to the document.")

	// ************** ADD TEXT TO A SPECIFIC TABLE CELL IN A GDOC FILE
	err = gdocsHelper.AddTextToTableCell(ctx, authConfig, copiedFile.Id, 0, 1, 1, "Cell Text")
	if err != nil {
		log.Fatalf("Unable to add text to table cell: %v", err)
	}
	fmt.Println("Added text to table cell.")

	// ************** SET COLOR TO A TABLE CELL
	color := &docs.OptionalColor{
		Color: &docs.Color{
			RgbColor: &docs.RgbColor{
				Red:   1.0, // Red color
				Green: 0.0,
				Blue:  0.0,
			},
		},
	}
	err = gdocsHelper.SetColorToTableCell(ctx, authConfig, copiedFile.Id, 0, 1, 1, color)
	if err != nil {
		log.Fatalf("Unable to set color to table cell: %v", err)
	}
	fmt.Println("Set color to table cell.")

	// ************** INSERT TEXT WITH URL
	// Step 3: Insert the URL of the original document into the copied document
	// First, get the end index of the copied document to insert at the end
	endIndex, err := gdocsHelper.GetDocumentEndIndex(ctx, authConfig, copiedFile.Id)
	if err != nil {
		log.Fatalf("main: unable to get end index of copied document: %v", err)
	}

	// Prepare the text and URL to insert
	textToInsert := gdocsHelper.GetDocumentURL(doc.DocumentId)
	urlToInsert := gdocsHelper.GetDocumentURL(doc.DocumentId)

	// Insert the text with link into the copied document
	err = gdocsHelper.InsertTextWithLinkAndRender(ctx, authConfig, copiedFile.Id, textToInsert, urlToInsert, endIndex-1)
	if err != nil {
		log.Fatalf("main: unable to insert text with link into copied document: %v", err)
	}

	fmt.Println("Inserted link to original document into copied document successfully.")
	fmt.Printf("You can view the copied document here: %s\n", gdocsHelper.GetDocumentURL(copiedFile.Id))

	// ************** ADD THE USER'S PERMISSION
	userEmail := "user@gmail.com"
	err = gDriveHelper.AddFolderPermission(ctx, authConfig, copiedFile.Id, userEmail, "reader")
	if err != nil {
		log.Fatalf("main: unable to add folder permission: %v", err)
	}

	fmt.Println("Added user's permission from folder successfully.")

	// ************** REMOVE THE USER'S PERMISSION
	userEmail = "user@gmail.com"
	err = gDriveHelper.RemoveFolderPermission(ctx, authConfig, copiedFile.Id, userEmail)
	if err != nil {
		log.Fatalf("main: unable to remove folder permission: %v", err)
	}

	fmt.Println("Removed user's permission from folder successfully.")

	// RENAME FOLDER GOOGLE DOC FILE
	newName := "New folder name"
	err = gDriveHelper.RenameFolder(ctx, authConfig, folder.Id, newName)
	if err != nil {
		log.Fatalf("main: unable to rename folder: %v", err)
	}

	fmt.Println("Folder renamed successfully.")

	// DELETE A GDOC FILE
	err = gDriveHelper.DeleteFileOrFolder(ctx, authConfig, doc.DocumentId)
	if err != nil {
		log.Fatalf("main: unable to delete file or folder: %v", err)
	}

	fmt.Println("File or folder deleted successfully.")
}
