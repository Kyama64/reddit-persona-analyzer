# Reddit Persona Analyzer

A Python tool that analyzes a Reddit user's profile to create a detailed persona based on their comments and posts, with export to Excel and Google Sheets.

## Features

- Analyzes a Reddit user's comments and posts
- Performs sentiment analysis on user's content
- Identifies most common words and topics
- Shows activity levels and engagement
- Identifies most active subreddits
- Generates a summary persona with emoji visualization
- Exports analysis to professionally formatted Excel (.xlsx) files
- Optional Google Sheets export (requires setup)

## Prerequisites

- Python 3.6+
- Reddit API credentials (see setup below)
- Google Cloud Project with Google Sheets API enabled (optional, for Google Sheets export)

## Setup

1. **Get Reddit API Credentials**
   - Go to https://www.reddit.com/prefs/apps/
   - Click "Create App" or "Create Another App" at the bottom
   - Select "script" as the app type
   - Fill in the name and description (can be anything)
   - Set the redirect URI to `http://localhost:8080`
   - Click "create app"
   - Note down the client ID (under the app name) and client secret

2. **Set up environment variables**
   - Create a `.env` file in the project directory with:
     ```
     REDDIT_CLIENT_ID=your_client_id_here
     REDDIT_CLIENT_SECRET=your_client_secret_here
     # Optional: Path to Google service account credentials JSON file
     GOOGLE_CREDENTIALS_PATH=credentials.json
     ```
   - Replace `your_client_id_here` and `your_client_secret_here` with your actual Reddit API credentials

3. **Set up Google Sheets API (Optional)**
   - Go to the [Google Cloud Console](https://console.cloud.google.com/)
   - Create a new project or select an existing one
   - Enable the Google Sheets API
   - Create a service account and download the JSON key file
   - Save the JSON file as `credentials.json` in the project directory
   - Share any Google Sheet you want to edit with the service account email address (found in the JSON file)

3. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### Basic Usage
Run the script:
```bash
python reddit_persona.py
```

When prompted, enter a Reddit profile URL (e.g., `https://www.reddit.com/user/username/` or `https://old.reddit.com/user/username/`) or just the username.

### Command Line Arguments
- `username`: The Reddit username to analyze (positional argument, required)
python reddit_persona.py username
```

### Export Options

Export to Excel (.xlsx) file:
```bash
python reddit_persona.py username --excel
```

The Excel file will be saved in the `exports` directory with a timestamped filename.

Export to Google Sheets (requires setup):
```bash
python reddit_persona.py username --export
```

Specify a Google Sheet ID:
```bash
python reddit_persona.py username --export --spreadsheet-id your_sheet_id_here
```

### Multiple Usernames

You can analyze multiple usernames by separating them with spaces:
```bash
python reddit_persona.py username1 username2 username3
```

## Example Output

The script will provide a detailed analysis including:
- Account information (creation date, karma)
- Comment analysis (sentiment, common words, active subreddits)
- Post analysis (top posts, active subreddits)
- A generated persona summary

## Notes

- The analysis is based on the most recent 100 comments and posts by default
- The tool respects Reddit's API rate limits
- Only public information is analyzed
- Some user profiles may be private or restricted

## Dependencies

- `praw` - Python Reddit API Wrapper
- `nltk` - Natural Language Toolkit for text processing
- `python-dotenv` - Loads environment variables from .env file
- `pandas` - Data manipulation and analysis
- `openpyxl` - Excel file support (used by pandas)
- `google-api-python-client` - Google API Client Library for Python (for Google Sheets export, optional)
- `google-auth-httplib2` - Google Authentication Library (for Google Sheets export, optional)
- `google-auth-oauthlib` - Google OAuth Library (for Google Sheets export, optional)

## Output

The tool will display the analysis in the console and save it to a professionally formatted Excel (.xlsx) file in the `exports` directory if the `--excel` flag is used. The output includes:

- Basic user information
- Personality and archetype analysis
- Sentiment analysis
- Motivations and goals
- Behaviors and habits
- Activity summary
- Most active subreddits

Excel files include:
- Proper formatting and styling
- Color-coded sections
- Auto-adjusted column widths
- Text wrapping for better readability

To enable Google Sheets export:
1. Set up the Google Sheets API as described in the Prerequisites section
2. Place the credentials file in the project directory
3. Use the `--export` flag when running the script
