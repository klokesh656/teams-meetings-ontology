# Microsoft Graph Transcript Extractor

## Project Overview
This is a Python script that extracts Microsoft Teams meeting transcripts using the Microsoft Graph API and stores them locally.

## Current Status
- [x] Create copilot-instructions.md file
- [x] Scaffold Python project structure
- [x] Create main script and configuration files
- [x] Install dependencies (Note: Windows Long Path support may be needed)
- [x] Create documentation

## Project Structure
- `src/transcript_extractor.py` - Main script for extracting transcripts
- `.env.example` - Template for environment variables
- `requirements.txt` - Python dependencies
- `transcripts/` - Output directory for downloaded transcripts
- `README.md` - Complete setup and usage documentation

## Development Guidelines
- Use Microsoft Graph SDK for Python
- Store credentials securely using environment variables
- Extract transcripts with OnlineMeetings.Transcript.Read.All permission
- Save transcripts to local folder with organized structure
- Handle Windows Long Path limitations on Windows systems

## Next Steps
1. Copy `.env.example` to `.env` and configure your Azure AD app credentials
2. Enable Windows Long Path support (if on Windows)
3. Install dependencies: `pip install msgraph-sdk azure-identity msal python-dotenv`
4. Run the script: `python src/transcript_extractor.py`
