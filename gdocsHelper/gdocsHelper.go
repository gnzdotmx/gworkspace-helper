package gdocsHelper

import (
	"context"
	"fmt"
	"io"
	"strings"

	"github.com/gnzdotmx/gworkspace-helper/auth"
	"google.golang.org/api/docs/v1"
	"google.golang.org/api/drive/v3"
	"google.golang.org/api/option"
)

// CreateGoogleDoc creates a new Google Doc with the given title.
func CreateGoogleDoc(ctx context.Context, config auth.Config, title string) (*docs.Document, error) {
	conf, token, err := auth.GetClient(ctx, config)
	if err != nil {
		return nil, fmt.Errorf("gdocsHelper: failed to get authenticated client: %w", err)
	}

	client := conf.Client(ctx, token)
	docsService, err := docs.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		return nil, fmt.Errorf("gdocsHelper: unable to create docs service: %w", err)
	}

	doc := &docs.Document{Title: title}
	createdDoc, err := docsService.Documents.Create(doc).Do()
	if err != nil {
		return nil, fmt.Errorf("gdocsHelper: unable to create document: %w", err)
	}
	return createdDoc, nil
}

// AddText appends text to the end of the document.
func AddText(ctx context.Context, config auth.Config, docID, text string) error {
	conf, token, err := auth.GetClient(ctx, config)
	if err != nil {
		return fmt.Errorf("gdocsHelper: failed to get authenticated client: %w", err)
	}

	client := conf.Client(ctx, token)
	docsService, err := docs.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		return fmt.Errorf("gdocsHelper: unable to create docs service: %w", err)
	}

	// Get the document to find the end index
	doc, err := docsService.Documents.Get(docID).Do()
	if err != nil {
		return fmt.Errorf("gdocsHelper: unable to retrieve document: %w", err)
	}

	endIndex := doc.Body.Content[len(doc.Body.Content)-1].EndIndex - 1

	requests := []*docs.Request{
		{
			InsertText: &docs.InsertTextRequest{
				Text: text,
				Location: &docs.Location{
					Index: endIndex,
				},
			},
		},
	}

	_, err = docsService.Documents.BatchUpdate(docID, &docs.BatchUpdateDocumentRequest{
		Requests: requests,
	}).Do()
	if err != nil {
		return fmt.Errorf("gdocsHelper: unable to add text to document: %w", err)
	}
	return nil
}

// ReplaceText replaces all occurrences of oldText with newText in the document.
func ReplaceText(ctx context.Context, config auth.Config, docID, oldText, newText string) error {
	conf, token, err := auth.GetClient(ctx, config)
	if err != nil {
		return fmt.Errorf("gdocsHelper: failed to get authenticated client: %w", err)
	}

	client := conf.Client(ctx, token)
	docsService, err := docs.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		return fmt.Errorf("gdocsHelper: unable to create docs service: %w", err)
	}

	requests := []*docs.Request{
		{
			ReplaceAllText: &docs.ReplaceAllTextRequest{
				ContainsText: &docs.SubstringMatchCriteria{
					Text:      oldText,
					MatchCase: true,
				},
				ReplaceText: newText,
			},
		},
	}

	_, err = docsService.Documents.BatchUpdate(docID, &docs.BatchUpdateDocumentRequest{
		Requests: requests,
	}).Do()
	if err != nil {
		return fmt.Errorf("gdocsHelper: unable to replace text in document: %w", err)
	}
	return nil
}

// MakeCopyOfGoogleDoc makes a copy of an existing Google Doc.
func MakeCopyOfGoogleDoc(ctx context.Context, config auth.Config, fileID, newTitle string) (*drive.File, error) {
	conf, token, err := auth.GetClient(ctx, config)
	if err != nil {
		return nil, fmt.Errorf("gdocsHelper: failed to get authenticated client: %w", err)
	}

	client := conf.Client(ctx, token)
	driveService, err := drive.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		return nil, fmt.Errorf("gdocsHelper: unable to create drive service: %w", err)
	}

	copiedFile := &drive.File{
		Name: newTitle,
	}

	file, err := driveService.Files.Copy(fileID, copiedFile).SupportsAllDrives(true).Do()
	if err != nil {
		return nil, fmt.Errorf("gdocsHelper: unable to copy file: %w", err)
	}
	return file, nil
}

// ExportGoogleDocAsText exports a Google Doc as plain text.
func ExportGoogleDocAsText(ctx context.Context, config auth.Config, fileID string) (string, error) {
	conf, token, err := auth.GetClient(ctx, config)
	if err != nil {
		return "", fmt.Errorf("gdocsHelper: failed to get authenticated client: %w", err)
	}

	client := conf.Client(ctx, token)
	driveService, err := drive.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		return "", fmt.Errorf("gdocsHelper: unable to create drive service: %w", err)
	}

	response, err := driveService.Files.Export(fileID, "text/plain").Download()
	if err != nil {
		return "", fmt.Errorf("gdocsHelper: unable to export file: %w", err)
	}
	defer response.Body.Close()

	content, err := io.ReadAll(response.Body)
	if err != nil {
		return "", fmt.Errorf("gdocsHelper: unable to read exported content: %w", err)
	}

	return string(content), nil
}

// RenameGoogleDoc renames a Google Doc.
func RenameGoogleDoc(ctx context.Context, config auth.Config, fileID, newTitle string) error {
	conf, token, err := auth.GetClient(ctx, config)
	if err != nil {
		return fmt.Errorf("gdocsHelper: failed to get authenticated client: %w", err)
	}

	client := conf.Client(ctx, token)
	driveService, err := drive.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		return fmt.Errorf("gdocsHelper: unable to create drive service: %w", err)
	}

	file := &drive.File{
		Name: newTitle,
	}

	_, err = driveService.Files.Update(fileID, file).Do()
	if err != nil {
		return fmt.Errorf("gdocsHelper: unable to rename file: %w", err)
	}
	return nil
}

// AddTextBetweenLines adds text between two known lines.
func AddTextBetweenLines(ctx context.Context, config auth.Config, docID, startLine, endLine, textToAdd string) error {
	conf, token, err := auth.GetClient(ctx, config)
	if err != nil {
		return fmt.Errorf("gdocsHelper: failed to get authenticated client: %w", err)
	}

	client := conf.Client(ctx, token)
	docsService, err := docs.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		return fmt.Errorf("gdocsHelper: unable to create docs service: %w", err)
	}

	// Retrieve the document
	doc, err := docsService.Documents.Get(docID).Do()
	if err != nil {
		return fmt.Errorf("gdocsHelper: unable to retrieve document: %w", err)
	}

	var startIndex, endIndex int64 = -1, -1

	// Find the start and end indices
	for _, element := range doc.Body.Content {
		if element.Paragraph != nil {
			for _, element := range element.Paragraph.Elements {
				textRun := element.TextRun
				if textRun != nil && textRun.Content != "" {
					content := strings.TrimSpace(textRun.Content)
					if content == startLine && startIndex == -1 {
						startIndex = element.StartIndex
					} else if content == endLine && endIndex == -1 {
						endIndex = element.StartIndex
					}
				}
			}
		}
	}

	if startIndex == -1 {
		return fmt.Errorf("gdocsHelper: start line '%s' not found", startLine)
	}
	if endIndex == -1 {
		return fmt.Errorf("gdocsHelper: end line '%s' not found", endLine)
	}
	if startIndex >= endIndex {
		return fmt.Errorf("gdocsHelper: start line occurs after end line")
	}

	// Insert text at the position after the start line
	insertIndex := startIndex + int64(len(startLine)) + 1 // +1 for newline

	requests := []*docs.Request{
		{
			InsertText: &docs.InsertTextRequest{
				Text: textToAdd + "\n",
				Location: &docs.Location{
					Index: insertIndex,
				},
			},
		},
	}

	_, err = docsService.Documents.BatchUpdate(docID, &docs.BatchUpdateDocumentRequest{
		Requests: requests,
	}).Do()
	if err != nil {
		return fmt.Errorf("gdocsHelper: unable to insert text between lines: %w", err)
	}

	return nil
}

// AddTextAfterLine adds text after a known line.
func AddTextAfterLine(ctx context.Context, config auth.Config, docID, lineContent, textToAdd string) error {
	conf, token, err := auth.GetClient(ctx, config)
	if err != nil {
		return fmt.Errorf("gdocsHelper: failed to get authenticated client: %w", err)
	}

	client := conf.Client(ctx, token)
	docsService, err := docs.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		return fmt.Errorf("gdocsHelper: unable to create docs service: %w", err)
	}

	// Retrieve the document
	doc, err := docsService.Documents.Get(docID).Do()
	if err != nil {
		return fmt.Errorf("gdocsHelper: unable to retrieve document: %w", err)
	}

	var insertIndex int64 = -1

	// Find the index after the known line
	for _, element := range doc.Body.Content {
		if element.Paragraph != nil {
			for _, elem := range element.Paragraph.Elements {
				textRun := elem.TextRun
				if textRun != nil && textRun.Content != "" {
					content := strings.TrimSpace(textRun.Content)
					if content == lineContent {
						insertIndex = elem.EndIndex
						break
					}
				}
			}
		}
	}

	if insertIndex == -1 {
		return fmt.Errorf("gdocsHelper: line '%s' not found", lineContent)
	}

	// Insert text after the known line
	requests := []*docs.Request{
		{
			InsertText: &docs.InsertTextRequest{
				Text: textToAdd + "\n",
				Location: &docs.Location{
					Index: insertIndex,
				},
			},
		},
	}

	_, err = docsService.Documents.BatchUpdate(docID, &docs.BatchUpdateDocumentRequest{
		Requests: requests,
	}).Do()
	if err != nil {
		return fmt.Errorf("gdocsHelper: unable to insert text after line: %w", err)
	}

	return nil
}

// AddTextAfterPatternInLine adds text after a known pattern in a line.
func AddTextAfterPatternInLine(ctx context.Context, config auth.Config, docID, pattern, textToAdd string) error {
	conf, token, err := auth.GetClient(ctx, config)
	if err != nil {
		return fmt.Errorf("gdocsHelper: failed to get authenticated client: %w", err)
	}

	client := conf.Client(ctx, token)
	docsService, err := docs.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		return fmt.Errorf("gdocsHelper: unable to create docs service: %w", err)
	}

	// Retrieve the document
	doc, err := docsService.Documents.Get(docID).Do()
	if err != nil {
		return fmt.Errorf("gdocsHelper: unable to retrieve document: %w", err)
	}

	var insertIndex int64 = -1

	// Find the pattern in the document
	for _, element := range doc.Body.Content {
		if element.Paragraph != nil {
			for _, elem := range element.Paragraph.Elements {
				textRun := elem.TextRun
				if textRun != nil && textRun.Content != "" {
					if idx := strings.Index(textRun.Content, pattern); idx != -1 {
						insertIndex = elem.StartIndex + int64(idx+len(pattern))
						break
					}
				}
			}
		}
	}

	if insertIndex == -1 {
		return fmt.Errorf("gdocsHelper: pattern '%s' not found", pattern)
	}

	// Insert text after the pattern
	requests := []*docs.Request{
		{
			InsertText: &docs.InsertTextRequest{
				Text: textToAdd,
				Location: &docs.Location{
					Index: insertIndex,
				},
			},
		},
	}

	_, err = docsService.Documents.BatchUpdate(docID, &docs.BatchUpdateDocumentRequest{
		Requests: requests,
	}).Do()
	if err != nil {
		return fmt.Errorf("gdocsHelper: unable to insert text after pattern: %w", err)
	}

	return nil
}

// AddTable adds a table to the Google Doc.
func AddTable(ctx context.Context, config auth.Config, docID string, rows, columns int64) error {
	conf, token, err := auth.GetClient(ctx, config)
	if err != nil {
		return fmt.Errorf("gdocsHelper: failed to get authenticated client: %w", err)
	}

	client := conf.Client(ctx, token)
	docsService, err := docs.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		return fmt.Errorf("gdocsHelper: unable to create docs service: %w", err)
	}

	// Get the document to find the insertion index
	doc, err := docsService.Documents.Get(docID).Do()
	if err != nil {
		return fmt.Errorf("gdocsHelper: unable to retrieve document: %w", err)
	}

	// Insert at the end of the document
	endIndex := doc.Body.Content[len(doc.Body.Content)-1].EndIndex - 1

	requests := []*docs.Request{
		{
			InsertTable: &docs.InsertTableRequest{
				Rows:    rows,
				Columns: columns,
				Location: &docs.Location{
					Index: endIndex,
				},
			},
		},
	}

	_, err = docsService.Documents.BatchUpdate(docID, &docs.BatchUpdateDocumentRequest{
		Requests: requests,
	}).Do()
	if err != nil {
		return fmt.Errorf("gdocsHelper: unable to add table: %w", err)
	}

	return nil
}

// AddTextToTableCell adds text to a specific cell in a table.
func AddTextToTableCell(ctx context.Context, config auth.Config, docID string, tableIndex, rowIndex, columnIndex int64, text string) error {
	conf, token, err := auth.GetClient(ctx, config)
	if err != nil {
		return fmt.Errorf("gdocsHelper: failed to get authenticated client: %w", err)
	}

	client := conf.Client(ctx, token)
	docsService, err := docs.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		return fmt.Errorf("gdocsHelper: unable to create docs service: %w", err)
	}

	// Retrieve the document
	doc, err := docsService.Documents.Get(docID).Do()
	if err != nil {
		return fmt.Errorf("gdocsHelper: unable to retrieve document: %w", err)
	}

	// Find the table
	var tableContent *docs.StructuralElement
	tableCount := int64(0)
	for _, content := range doc.Body.Content {
		if content.Table != nil {
			if tableCount == tableIndex {
				tableContent = content
				break
			}
			tableCount++
		}
	}

	if tableContent == nil {
		return fmt.Errorf("gdocsHelper: table at index %d not found", tableIndex)
	}

	// Get the cell's start index
	cell := tableContent.Table.TableRows[rowIndex].TableCells[columnIndex]
	insertIndex := cell.StartIndex + 1 // +1 to go inside the cell

	requests := []*docs.Request{
		{
			InsertText: &docs.InsertTextRequest{
				Text: text,
				Location: &docs.Location{
					Index: insertIndex,
				},
			},
		},
	}

	_, err = docsService.Documents.BatchUpdate(docID, &docs.BatchUpdateDocumentRequest{
		Requests: requests,
	}).Do()
	if err != nil {
		return fmt.Errorf("gdocsHelper: unable to add text to table cell: %w", err)
	}

	return nil
}

// AddLinkToText adds a hyperlink to specific text in the document.
func AddLinkToText(ctx context.Context, config auth.Config, docID, searchText, url string) error {
	conf, token, err := auth.GetClient(ctx, config)
	if err != nil {
		return fmt.Errorf("gdocsHelper: failed to get authenticated client: %w", err)
	}

	client := conf.Client(ctx, token)
	docsService, err := docs.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		return fmt.Errorf("gdocsHelper: unable to create docs service: %w", err)
	}

	// Find the text in the document
	doc, err := docsService.Documents.Get(docID).Do()
	if err != nil {
		return fmt.Errorf("gdocsHelper: unable to retrieve document: %w", err)
	}

	var startIndex, endIndex int64 = -1, -1

	for _, element := range doc.Body.Content {
		if element.Paragraph != nil {
			for _, elem := range element.Paragraph.Elements {
				textRun := elem.TextRun
				if textRun != nil && textRun.Content != "" {
					idx := strings.Index(textRun.Content, searchText)
					if idx != -1 {
						startIndex = elem.StartIndex + int64(idx)
						endIndex = startIndex + int64(len(searchText))
						break
					}
				}
			}
		}
	}

	if startIndex == -1 || endIndex == -1 {
		return fmt.Errorf("gdocsHelper: text '%s' not found", searchText)
	}

	// Update the text style to include the link
	requests := []*docs.Request{
		{
			UpdateTextStyle: &docs.UpdateTextStyleRequest{
				Fields: "link",
				Range: &docs.Range{
					StartIndex: startIndex,
					EndIndex:   endIndex,
				},
				TextStyle: &docs.TextStyle{
					Link: &docs.Link{
						Url: url,
					},
				},
			},
		},
	}

	_, err = docsService.Documents.BatchUpdate(docID, &docs.BatchUpdateDocumentRequest{
		Requests: requests,
	}).Do()
	if err != nil {
		return fmt.Errorf("gdocsHelper: unable to add link to text: %w", err)
	}

	return nil
}

// ReplaceMultipleTexts replaces multiple strings in the Google Doc.
func ReplaceMultipleTexts(ctx context.Context, config auth.Config, docID string, replacements map[string]string) error {
	conf, token, err := auth.GetClient(ctx, config)
	if err != nil {
		return fmt.Errorf("gdocsHelper: failed to get authenticated client: %w", err)
	}

	client := conf.Client(ctx, token)
	docsService, err := docs.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		return fmt.Errorf("gdocsHelper: unable to create docs service: %w", err)
	}

	var requests []*docs.Request

	for oldText, newText := range replacements {
		requests = append(requests, &docs.Request{
			ReplaceAllText: &docs.ReplaceAllTextRequest{
				ContainsText: &docs.SubstringMatchCriteria{
					Text:      oldText,
					MatchCase: true,
				},
				ReplaceText: newText,
			},
		})
	}

	_, err = docsService.Documents.BatchUpdate(docID, &docs.BatchUpdateDocumentRequest{
		Requests: requests,
	}).Do()

	if err != nil {
		return fmt.Errorf("gdocsHelper: unable to replace multiple texts: %w", err)
	}

	return nil
}

// SetColorToTableCell sets the background color of a specific table cell.
func SetColorToTableCell(ctx context.Context, config auth.Config, docID string, tableIndex, rowIndex, columnIndex int64, color *docs.OptionalColor) error {
	conf, token, err := auth.GetClient(ctx, config)
	if err != nil {
		return fmt.Errorf("gdocsHelper: failed to get authenticated client: %w", err)
	}

	client := conf.Client(ctx, token)
	docsService, err := docs.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		return fmt.Errorf("gdocsHelper: unable to create docs service: %w", err)
	}

	// Retrieve the document
	doc, err := docsService.Documents.Get(docID).Do()
	if err != nil {
		return fmt.Errorf("gdocsHelper: unable to retrieve document: %w", err)
	}

	// Find the table
	var tableContent *docs.StructuralElement
	tableCount := int64(0)
	for _, content := range doc.Body.Content {
		if content.Table != nil {
			if tableCount == tableIndex {
				tableContent = content
				break
			}
			tableCount++
		}
	}

	if tableContent == nil {
		return fmt.Errorf("gdocsHelper: table at index %d not found", tableIndex)
	}

	// Update the cell style
	requests := []*docs.Request{
		{
			UpdateTableCellStyle: &docs.UpdateTableCellStyleRequest{
				TableCellStyle: &docs.TableCellStyle{
					BackgroundColor: color,
				},
				Fields: "backgroundColor",
				TableRange: &docs.TableRange{
					TableCellLocation: &docs.TableCellLocation{
						TableStartLocation: &docs.Location{
							Index: tableContent.StartIndex,
						},
						RowIndex:    rowIndex,
						ColumnIndex: columnIndex,
					},
					RowSpan:    1,
					ColumnSpan: 1,
				},
			},
		},
	}

	_, err = docsService.Documents.BatchUpdate(docID, &docs.BatchUpdateDocumentRequest{
		Requests: requests,
	}).Do()

	if err != nil {
		return fmt.Errorf("gdocsHelper: unable to set color to table cell: %w", err)
	}

	return nil
}

// AddFilePermission adds permissions to a file for a specific email.
func AddFilePermission(ctx context.Context, config auth.Config, fileID, email, role string) error {
	conf, token, err := auth.GetClient(ctx, config)
	if err != nil {
		return fmt.Errorf("gdocsHelper: failed to get authenticated client: %w", err)
	}

	client := conf.Client(ctx, token)
	driveService, err := drive.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		return fmt.Errorf("gdocsHelper: unable to create drive service: %w", err)
	}

	permission := &drive.Permission{
		Type:         "user",
		Role:         role, // e.g., "owner", "writer", "reader"
		EmailAddress: email,
	}

	_, err = driveService.Permissions.Create(fileID, permission).SendNotificationEmail(false).Do()
	if err != nil {
		return fmt.Errorf("gdocsHelper: unable to add permission to file: %w", err)
	}

	return nil
}

// InsertTextWithLinkAndRender inserts text into a Google Doc, applies a hyperlink to it,
// and ensures it's rendered properly.
func InsertTextWithLinkAndRender(ctx context.Context, config auth.Config, docID, text, url string, locationIndex int64) error {
	conf, token, err := auth.GetClient(ctx, config)
	if err != nil {
		return fmt.Errorf("gdocsHelper: failed to get authenticated client: %w", err)
	}

	client := conf.Client(ctx, token)
	docsService, err := docs.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		return fmt.Errorf("gdocsHelper: unable to create docs service: %w", err)
	}

	// Step 1: Insert the text at the specified location
	insertTextRequest := &docs.Request{
		InsertText: &docs.InsertTextRequest{
			Text: text,
			Location: &docs.Location{
				Index: locationIndex,
			},
		},
	}

	// Step 2: Apply the hyperlink style to the inserted text
	// We need to know the start and end indices of the inserted text
	textLength := int64(len(text))
	updateTextStyleRequest := &docs.Request{
		UpdateTextStyle: &docs.UpdateTextStyleRequest{
			Fields: "link",
			Range: &docs.Range{
				StartIndex: locationIndex,
				EndIndex:   locationIndex + textLength,
			},
			TextStyle: &docs.TextStyle{
				Link: &docs.Link{
					Url: url,
				},
			},
		},
	}

	// Batch the requests together
	requests := []*docs.Request{
		insertTextRequest,
		updateTextStyleRequest,
	}

	_, err = docsService.Documents.BatchUpdate(docID, &docs.BatchUpdateDocumentRequest{
		Requests: requests,
	}).Do()
	if err != nil {
		return fmt.Errorf("gdocsHelper: unable to insert text with link: %w", err)
	}

	return nil
}

// GetDocumentEndIndex retrieves the index at the end of the document's body content.
func GetDocumentEndIndex(ctx context.Context, config auth.Config, docID string) (int64, error) {
	conf, token, err := auth.GetClient(ctx, config)
	if err != nil {
		return 0, fmt.Errorf("gdocsHelper: failed to get authenticated client: %w", err)
	}

	client := conf.Client(ctx, token)
	docsService, err := docs.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		return 0, fmt.Errorf("gdocsHelper: unable to create docs service: %w", err)
	}

	doc, err := docsService.Documents.Get(docID).Do()
	if err != nil {
		return 0, fmt.Errorf("gdocsHelper: unable to retrieve document: %w", err)
	}

	// The end index of the last element in the body content
	endIndex := doc.Body.Content[len(doc.Body.Content)-1].EndIndex

	return endIndex, nil
}

// GetDocumentURL constructs the URL of a Google Doc given its file ID.
func GetDocumentURL(fileID string) string {
	return fmt.Sprintf("https://docs.google.com/document/d/%s/edit", fileID)
}
