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
