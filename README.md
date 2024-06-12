# PMAX Search Terms Script Enhanced - V1.0.1

This Google Ads script retrieves search term data for Performance Max campaigns and exports it to a Google Spreadsheet. It also sends an email notification when the script has run.

## Setup Instructions

1. Open the Google Ads account where you want to run the script.
2. Navigate to Tools & Settings > Bulk Actions > Scripts.
3. Click on the plus button to create a new script.
4. Copy and paste the script into the code editor.
5. Configure the following variables in the `config` object:
   - `LOG`: Set to `true` to enable logging, or `false` to disable (default: `false`).
   - `DATE_RANGE`: Specify the date range for retrieving search terms (e.g., `last_n_days(30)` for the last 30 days).
   - `SPREADSHEET_URL`: Enter the URL of the Google Spreadsheet where the data will be exported. Make sure to include only the URL up to `/edit`.
   - `EMAIL_ADDRESSES`: Enter the email addresses to receive the notification when the script has run. Separate multiple addresses with commas.
6. Save the script and authorize it to access the necessary Google services.

## Configuration

- `LOG`: Set to `true` to enable logging for debugging purposes. Default is `false`.
- `DATE_RANGE`: Specify the date range for retrieving search terms. Use the `last_n_days(n)` function to retrieve data for the last `n` days.
- `SPREADSHEET_URL`: Enter the URL of the Google Spreadsheet where the data will be exported. Only include the URL up to `/edit`.
- `EMAIL_ADDRESSES`: Enter the email addresses to receive the notification when the script has run. Separate multiple addresses with commas.

## Notes

- To run this script, make sure the permission of your Google Spreadsheet is set to 'Anyone with the link' and editor mode is enabled.
- This script does not work at the MCC (My Client Center) level.

## Functions

### `main()`

The main function that orchestrates the script execution. It retrieves search term data for enabled Performance Max campaigns, exports the data to a Google Spreadsheet, and sends an email notification.

### `checkTab(file)`

Creates a new sheet in the specified Google Spreadsheet based on the current date. If a sheet with the current date already exists, it selects that sheet. It also removes the default "Sheet1".

### `formatSheet(sheet)`

Applies formatting to the specified sheet, including setting the header row style, applying number formatting to data cells, and enabling row banding.

### `last_n_days(n)`

Calculates the date range for the last `n` days.

### `google_date_range(from, to)`

Formats the date range in the Google Ads API format.

## Version History

- V1.0.1: Initial release of the enhanced PMAX Search Terms script.
- V1.0.2: Added conditional formatting for conversion value metric.

For any issues or questions, please contact [your email address or support channel].
