# TODO

## Copilot Pages

- [ ] Restore `microsoft_list_copilot_pages`, `microsoft_preview_copilot_page`, and `microsoft_import_copilot_page` after SharePoint Embedded access is ready.
- [ ] Grant the MCP Entra application app-only guest access to the Loop/Copilot Pages owning application and verify the tenant-side configuration.
- [ ] Implement the app-only certificate/client-credential token flow required by the SharePoint Embedded container API.
- [ ] Bind app-only container reads to the signed-in user's own Copilot Pages containers; never expose another user's containers.
- [ ] Re-add the minimum required Pages Graph permissions and perform a fresh Microsoft OAuth consent after re-enabling the tools.
- [ ] Test Pages list, preview, and confirmed import end-to-end in MCP Inspector.

The Pages adapter methods remain in `src/standard-microsoft.mjs`, but the MCP registrations are commented out until these items are complete.
