package gMeetHelper

import (
	"context"
	"fmt"
	"time"

	"github.com/gnzdotmx/gworkspace-helper/auth"
	"google.golang.org/api/calendar/v3"
	"google.golang.org/api/drive/v3"
	"google.golang.org/api/option"
)

// CreateCalendarEvent creates a new event in Google Calendar.
func CreateCalendarEvent(ctx context.Context, config auth.Config, summary, location, description string, startTime, endTime time.Time) (*calendar.Event, error) {
	conf, token, err := auth.GetClient(ctx, config)
	if err != nil {
		return nil, fmt.Errorf("gMeetHelper: failed to get authenticated client: %w", err)
	}

	client := conf.Client(ctx, token)
	calendarService, err := calendar.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		return nil, fmt.Errorf("gMeetHelper: unable to create calendar service: %w", err)
	}

	// Load the Japanese timezone
	jst, err := time.LoadLocation("Asia/Tokyo")
	if err != nil {
		return nil, fmt.Errorf("gMeetHelper: unable to load Japanese timezone: %w", err)
	}

	// Convert startTime and endTime to Japanese timezone
	startTimeJST := startTime.In(jst)
	endTimeJST := endTime.In(jst)

	event := &calendar.Event{
		Summary:     summary,
		Location:    location,
		Description: description,
		Start: &calendar.EventDateTime{
			DateTime: startTimeJST.Format(time.RFC3339),
			TimeZone: "Asia/Tokyo",
		},
		End: &calendar.EventDateTime{
			DateTime: endTimeJST.Format(time.RFC3339),
			TimeZone: "Asia/Tokyo",
		},
		ConferenceData: &calendar.ConferenceData{
			CreateRequest: &calendar.CreateConferenceRequest{
				RequestId: "unique-request-id", // Must be unique for each request
				ConferenceSolutionKey: &calendar.ConferenceSolutionKey{
					Type: "hangoutsMeet",
				},
			},
		},
	}

	createdEvent, err := calendarService.Events.Insert("primary", event).ConferenceDataVersion(1).Do()
	if err != nil {
		return nil, fmt.Errorf("gMeetHelper: unable to create calendar event: %w", err)
	}
	return createdEvent, nil
}

// AddAttendeesToEvent adds attendees to an existing event.
func AddAttendeesToEvent(ctx context.Context, config auth.Config, eventID string, attendees []string) error {
	conf, token, err := auth.GetClient(ctx, config)
	if err != nil {
		return fmt.Errorf("gMeetHelper: failed to get authenticated client: %w", err)
	}

	client := conf.Client(ctx, token)
	calendarService, err := calendar.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		return fmt.Errorf("gMeetHelper: unable to create calendar service: %w", err)
	}

	event, err := calendarService.Events.Get("primary", eventID).Do()
	if err != nil {
		return fmt.Errorf("gMeetHelper: unable to retrieve event: %w", err)
	}

	for _, email := range attendees {
		event.Attendees = append(event.Attendees, &calendar.EventAttendee{Email: email})
	}

	_, err = calendarService.Events.Update("primary", event.Id, event).Do()
	if err != nil {
		return fmt.Errorf("gMeetHelper: unable to add attendees to event: %w", err)
	}
	return nil
}

// AttachFileToEvent attaches a file to an event.
func AttachFileToEvent(ctx context.Context, config auth.Config, eventID, fileID string) error {
	// Authenticate and create the client
	conf, token, err := auth.GetClient(ctx, config)
	if err != nil {
		return fmt.Errorf("gMeetHelper: failed to get authenticated client: %w", err)
	}
	client := conf.Client(ctx, token)

	// Create the Drive service to retrieve file metadata
	driveService, err := drive.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		return fmt.Errorf("gMeetHelper: unable to create Drive service: %w", err)
	}

	// Get the file metadata from Drive
	file, err := driveService.Files.Get(fileID).Fields("webViewLink", "name", "mimeType").Do()
	if err != nil {
		return fmt.Errorf("gMeetHelper: unable to retrieve file metadata: %w", err)
	}

	// Create the Calendar service
	calendarService, err := calendar.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		return fmt.Errorf("gMeetHelper: unable to create Calendar service: %w", err)
	}

	// Retrieve the event
	event, err := calendarService.Events.Get("primary", eventID).Do()
	if err != nil {
		return fmt.Errorf("gMeetHelper: unable to retrieve event: %w", err)
	}

	// Create the EventAttachment with required fields
	attachment := &calendar.EventAttachment{
		FileId:   fileID,
		FileUrl:  file.WebViewLink,
		Title:    file.Name,
		MimeType: file.MimeType,
	}

	// Initialize the Attachments slice if it's nil
	if event.Attachments == nil {
		event.Attachments = []*calendar.EventAttachment{}
	}

	// Append the attachment to the event
	event.Attachments = append(event.Attachments, attachment)

	// Update the event with supportsAttachments set to true
	_, err = calendarService.Events.Update("primary", event.Id, event).SupportsAttachments(true).Do()
	if err != nil {
		return fmt.Errorf("gMeetHelper: unable to attach file to event: %w", err)
	}

	return nil
}
