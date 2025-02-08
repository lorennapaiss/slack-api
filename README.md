# Slack Excel Notifier 📩🔔

## Description
This Node.js script reads an Excel file containing a list of responsible persons, their emails, and the indicators they manage. It then sends personalized messages via Slack, reminding them to update their indicators before the deadline. The script:

- Extracts and groups indicators by responsible person.
- Retrieves Slack user IDs based on emails.
- Sends customized messages via Slack DM using the Slack API.

## Features
✅ Reads an Excel file (.xlsx) and processes user data.  
✅ Groups indicators by responsible email.  
✅ Sends formatted messages using Slack's Markdown support.  
✅ Handles errors gracefully.  

## Usage

1. Update the `token` variable with your Slack bot token.  
2. Ensure the Excel file follows the expected format (columns: `Nome`, `Email`, `Indicador`).  
3. Run the script with:
   ```sh
   sendMessageToUsers.js
   ```

🚀 Easily automate Slack notifications for key reminders!
