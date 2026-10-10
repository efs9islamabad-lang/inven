# Upgrade notes

- Modernized responsive visual design and improved mobile layout.
- Added lock protection around stock write operations.
- Added server-side SKU checks and quantity validation.
- Stock removal now recalculates and validates the real server-side balance instead of trusting the browser value.
- Replaced short timestamp IDs with UUID-based references to reduce collisions.
- Fixed settings save behavior so unedited settings are not deleted.
- Added date range validation and local-time-safe default report dates.
- Added HTML escaping for SKU/supplier dropdowns and failure handlers for key Apps Script calls.
- Updated naming to avoid falsely implying this is an Excel VBA file; it is a Google Apps Script web app.

## Still required before commercial sale
- Run end-to-end tests in a duplicate Google Sheet with multiple users and realistic data.
- Implement product/supplier create-edit-delete screens if these are required; the current interface does not provide full master-data CRUD.
- Test sample-data loading on a disposable copy; it is destructive to existing sample areas and should not be run on live records.
- Review deployment permissions and privacy disclosures.
