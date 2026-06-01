# Lecture Registration LineBot Tracker

This project integrates **Google Forms**, **Google Apps Script (GAS)**, and the **LINE Messaging API** to create a seamless event registration tracking system. It automatically manages "Accepted" (正取) and "Waitlisted" (備取) attendees based on a capacity limit, and allows the organizer to query real-time registration status directly through a LINE Bot.

## 🌟 Features

*   **Automated Capacity Management:** Automatically assigns "Accepted" (正取) or "Waitlisted" (備取) to new registrants upon Google Form submission based on a customizable limit.
*   **Sequential Numbering:** Automatically generates and assigns an order number (序號) to each participant.
*   **LINE Bot Integration:** Instantly query current registration counts and waitlist status via LINE text messages.
*   **VIP Tracking:** Identify and list attendees who registered using a specific "VIP" invite code.
*   **Automated Push Notifications:** Sends periodic summary reports (every 12 hours) to a designated LINE user or group, including overall counts and the VIP list.

## 🛠️ Prerequisites

*   A Google Account (for Google Forms and Google Sheets).
*   A LINE Developers Account with a registered Messaging API Channel.

## 🚀 Setup & Installation

### 1. Google Form & Spreadsheet Setup
1. Create a Google Form with your desired questions (e.g., Name, Email, Phone, Dog Name, Invite Code).
2. Go to the "Responses" tab and link it to a Google Spreadsheet.
3. Open the linked Spreadsheet, click on **Extensions > Apps Script**.

### 2. Apps Script Configuration
1. Copy the code from `Code.gs` in this repository and paste it into your Apps Script editor.
2. Update the Global Variables in the code:
```javascript
   const LINE_ACCESS_TOKEN = 'YOUR_LINE_CHANNEL_ACCESS_TOKEN';
   const LINE_USER_ID = 'YOUR_LINE_USER_ID'; // For periodic push notifications
```
### 3. Update the capacity limit in handleFormSubmit() and recalcAllStatus() if needed:
```javascript
const LIMIT = 3; // Change this to your actual maximum capacity
```

3. Deploy the Webhook
- In the Apps Script editor, click Deploy > New deployment.
- Select Web app as the deployment type.
- Set Execute as to Me and Who has access to Anyone.
- Click Deploy and copy the generated Web app URL.
- Go to your LINE Developers Console, paste the URL into the Webhook URL setting of your Messaging API channel, and enable "Use webhook".

4. Initialize Triggers
Run the following functions manually once from the Apps Script editor (you will be prompted to grant permissions):
```javascript
- setupFormSubmitTrigger(): "Sets up the trigger to process new form submissions automatically."
- setupPeriodicTrigger(): "(Optional) Sets up the 12-hour automated push notification."
```

📱 LINE Bot Commands
Interact with your LINE Bot using the following keywords (case-insensitive):
| Keyword / Input | Bot Response |
| :--- | :--- |
| **人數**, **狀態**, 或 **正取** | Returns the current registration status, including the number of Accepted (✅) and Waitlisted (⏳) participants. Gives a prompt if the waitlist is getting long. |
| **VIP** | Returns a detailed list of all participants who registered using the "VIP" invite code, including their registration time, name, dog's name, and current status. |

🗂️ Data Structure Reference
The script automatically detects or creates the following required columns in your Google Sheet:
- 狀態 (Status): Automatically filled with "正取" or "備取".
- 序號 (Order): Automatically filled with the sequence number.
- 邀請碼 (Invite Code): Used to determine VIP status.

⚠️ Important Notes
- Do not publicly share your LINE_ACCESS_TOKEN. Keep it secure.
- If you manually edit the spreadsheet rows, you can run recalcAllStatus() from the Apps Script editor to recalculate and refresh all statuses and order numbers.
