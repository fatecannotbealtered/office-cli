package cmd

import (
	"fmt"

	"github.com/fatecannotbealtered/office-cli/internal/config"
	"github.com/fatecannotbealtered/office-cli/internal/output"
	"github.com/spf13/cobra"
)

var setupCmd = &cobra.Command{
	Use:   "setup",
	Short: "Create or verify the office-cli configuration file",
	Long: `Creates ~/.office-cli/config.json with sensible defaults if it does not exist.
If the file already exists, validates it and reports any issues.`,
	RunE: runSetup,
}

func init() {
	rootCmd.AddCommand(setupCmd)
}

func runSetup(_ *cobra.Command, _ []string) error {
	path := config.ConfigPath()

	// Try to load existing config
	issues, err := config.Validate()
	if err != nil {
		return emitError("failed to read config: "+err.Error(), output.ErrEngine, path, ExitEngine)
	}

	if len(issues) == 0 {
		cfg := config.Load()
		if jsonMode {
			output.PrintJSON(map[string]any{
				"status":  "ok",
				"path":    path,
				"mode":    cfg.Permissions.Mode,
				"message": "config is valid",
			})
			return nil
		}
		output.Success(fmt.Sprintf("config is valid: %s", path))
		output.Info(fmt.Sprintf("  permission mode: %s", cfg.Permissions.Mode))
		return nil
	}

	// Config has issues or doesn't exist — create/update it
	if err := config.Setup(); err != nil {
		return emitError("failed to create config: "+err.Error(), output.ErrPermission, path, ExitForbidden)
	}

	cfg := config.Load()
	if jsonMode {
		output.PrintJSON(map[string]any{
			"status": "created",
			"path":   path,
			"mode":   cfg.Permissions.Mode,
		})
		return nil
	}
	output.Success(fmt.Sprintf("config created: %s", path))
	output.Info(fmt.Sprintf("  permission mode: %s", cfg.Permissions.Mode))
	output.Info("  edit the file to change email, timezone, or permissions")
	return nil
}
