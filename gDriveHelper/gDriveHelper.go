package gDriveHelper

import (
	"context"
	"fmt"

	"github.com/gnzdotmx/gworkspace-helper/auth"
	"google.golang.org/api/drive/v3"
	"google.golang.org/api/option"
)

// CreateFolder creates a new folder in Google Drive.
func CreateFolder(ctx context.Context, config auth.Config, name string) (*drive.File, error) {
	conf, token, err := auth.GetClient(ctx, config)
	if err != nil {
		return nil, fmt.Errorf("gDriveHelper: failed to get authenticated client: %w", err)
	}

	client := conf.Client(ctx, token)
	driveService, err := drive.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		return nil, fmt.Errorf("gDriveHelper: unable to create drive service: %w", err)
	}

	folder := &drive.File{
		Name:     name,
		MimeType: "application/vnd.google-apps.folder",
	}

	createdFolder, err := driveService.Files.Create(folder).SupportsAllDrives(true).Do()
	if err != nil {
		return nil, fmt.Errorf("gDriveHelper: unable to create folder: %w", err)
	}
	return createdFolder, nil
}

// AddFolderPermission adds permissions to the folder for a specific email.
func AddFolderPermission(ctx context.Context, config auth.Config, folderID, email, role string) error {
	conf, token, err := auth.GetClient(ctx, config)
	if err != nil {
		return fmt.Errorf("gDriveHelper: failed to get authenticated client: %w", err)
	}

	client := conf.Client(ctx, token)
	driveService, err := drive.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		return fmt.Errorf("gDriveHelper: unable to create drive service: %w", err)
	}

	permission := &drive.Permission{
		Type:         "user",
		Role:         role, // e.g., "owner", "writer", "reader"
		EmailAddress: email,
	}

	_, err = driveService.Permissions.Create(folderID, permission).SupportsAllDrives(true).Do()
	if err != nil {
		return fmt.Errorf("gDriveHelper: unable to add permission to folder: %w", err)
	}
	return nil
}

// CopyFileToFolder copies a file to the specified folder.
func CopyFileToFolder(ctx context.Context, config auth.Config, fileID, folderID string) (*drive.File, error) {
	conf, token, err := auth.GetClient(ctx, config)
	if err != nil {
		return nil, fmt.Errorf("gDriveHelper: failed to get authenticated client: %w", err)
	}

	client := conf.Client(ctx, token)
	driveService, err := drive.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		return nil, fmt.Errorf("gDriveHelper: unable to create drive service: %w", err)
	}

	file, err := driveService.Files.Copy(fileID, &drive.File{
		Parents: []string{folderID},
	}).SupportsAllDrives(true).Do()
	if err != nil {
		return nil, fmt.Errorf("gDriveHelper: unable to copy file to folder: %w", err)
	}
	return file, nil
}

// RemoveFolderPermission removes a permission from a folder for a specific user.
func RemoveFolderPermission(ctx context.Context, config auth.Config, folderID, email string) error {
	conf, token, err := auth.GetClient(ctx, config)
	if err != nil {
		return fmt.Errorf("gDriveHelper: failed to get authenticated client: %w", err)
	}

	client := conf.Client(ctx, token)
	driveService, err := drive.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		return fmt.Errorf("gDriveHelper: unable to create drive service: %w", err)
	}

	// List permissions to find the permission ID for the given email
	permissionsList, err := driveService.Permissions.List(folderID).Fields("permissions(id,emailAddress)").SupportsAllDrives(true).Do()
	if err != nil {
		return fmt.Errorf("gDriveHelper: unable to list permissions for folder: %w", err)
	}

	var permissionID string
	for _, permission := range permissionsList.Permissions {
		if permission.EmailAddress == email {
			permissionID = permission.Id
			break
		}
	}

	if permissionID == "" {
		return fmt.Errorf("gDriveHelper: no permission found for email %s", email)
	}

	// Delete the permission
	err = driveService.Permissions.Delete(folderID, permissionID).SupportsAllDrives(true).Do()
	if err != nil {
		return fmt.Errorf("gDriveHelper: unable to remove permission: %w", err)
	}

	return nil
}

// RenameFolder renames a folder in Google Drive.
func RenameFolder(ctx context.Context, config auth.Config, folderID, newName string) error {
	conf, token, err := auth.GetClient(ctx, config)
	if err != nil {
		return fmt.Errorf("gDriveHelper: failed to get authenticated client: %w", err)
	}

	client := conf.Client(ctx, token)
	driveService, err := drive.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		return fmt.Errorf("gDriveHelper: unable to create drive service: %w", err)
	}

	folder := &drive.File{
		Name: newName,
	}

	_, err = driveService.Files.Update(folderID, folder).SupportsAllDrives(true).Do()
	if err != nil {
		return fmt.Errorf("gDriveHelper: unable to rename folder: %w", err)
	}

	return nil
}

// DeleteFolder deletes a file or folder from Google Drive.
func DeleteFileOrFolder(ctx context.Context, config auth.Config, folderFileID string) error {
	conf, token, err := auth.GetClient(ctx, config)
	if err != nil {
		return fmt.Errorf("gDriveHelper: failed to get authenticated client: %w", err)
	}

	client := conf.Client(ctx, token)
	driveService, err := drive.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		return fmt.Errorf("gDriveHelper: unable to create drive service: %w", err)
	}

	err = driveService.Files.Delete(folderFileID).SupportsAllDrives(true).Do()
	if err != nil {
		return fmt.Errorf("gDriveHelper: unable to delete folder or file: %w", err)
	}

	return nil
}
