<p align="center">
  <a href="https://github.com/Pulkit1822/googleScript-myday-scheduler">
    <img src="https://upload.wikimedia.org/wikipedia/commons/thumb/2/2f/Google_Apps_Script.svg/1024px-Google_Apps_Script.svg.png" height="128">
  </a>
  <h1 align="center">MyDay</h1>
</p>

## Description

MyDay is a Google Sheets-based project designed to help you organize your daily tasks and assignments efficiently. It provides a structured layout for tracking tasks and assignments, setting deadlines, and monitoring their status. The project leverages Google Apps Script to automate the creation and management of your daily schedule.

## Installation

1. Open Google Sheets.
2. Go to `Extensions` > `Apps Script`.
3. Copy and paste the code into the Apps Script editor.
4. Save the project.

## Usage

1. Open the Google Sheet where you installed the script.
2. Click on the `myday` menu.
3. Select `Add New Day's Schedule` to create a new schedule for the day.
4. Fill in your tasks and assignments, set deadlines, and update their status as needed.

## Technologies Used

- **Google Apps Script**: Automates the creation and management of the daily schedule.
- **Google Sheets**: Provides a familiar and user-friendly interface for managing tasks and assignments.
- **Conditional Formatting**: Highlights the status of tasks and assignments for easy tracking.

## Benefits

- **Automation**: Reduces manual effort in setting up and managing daily schedules.
- **Customization**: Allows users to tailor the schedule to their specific needs.
- **Efficiency**: Streamlines the process of tracking tasks and assignments, making it easier to stay organized.

## Comparison with Similar Projects

Unlike other task management tools, MyDay is fully integrated with Google Sheets, making it accessible and easy to use for anyone familiar with spreadsheets. It leverages the power of Google Apps Script to provide automation and customization that other tools may lack.

## Running the Project

1. Open Google Sheets.
2. Go to `Extensions` > `Apps Script`.
3. Copy and paste the code into the Apps Script editor.
4. Save the project.
5. Open the Google Sheet where you installed the script.
6. Click on the `myday` menu.
7. Select `Add New Day's Schedule` to create a new schedule for the day.
8. Fill in your tasks and assignments, set deadlines, and update their status as needed.
9. To sync events with Google Calendar, select `Sync All Events to Calendar` from the `myday` menu.
10. To initialize calendar access, select `Initialize Calendar Access` from the `myday` menu.

### Note

When running the script for the first time, you will be prompted to grant necessary permissions for the script to access Google Calendar and Google Sheets. Make sure to review and grant these permissions to ensure the script functions correctly.

### Adding the `appsscript.json` File

1. In the Apps Script editor, click on the `+` button next to the `Files` section.
2. Select `Script Properties` and then click on `appsscript.json`.
3. Copy the contents of the `appsscript.json` file provided in this repository.
4. Paste the contents into the `appsscript.json` file in the Apps Script editor.
5. Save the project.

### Adding the Calendar Service

1. In the Apps Script editor, click on `Services` (+ icon) in the left sidebar.
2. Click `Add a service` (+ button).
3. Find and select `Google Calendar API`.
4. Click `Add`.

### Clearing Previous Authorization Cache and Reviewing Permissions

1. In the Apps Script editor, click on `View` → `Show project manifest`.
2. If `appsscript.json` exists, replace its content with the code above.
3. If it doesn't exist, create it by clicking the `+` next to Files and choosing `Script`, name it `appsscript.json`.
4. Click `Review Permissions` in the editor.
5. At the top right, click on your Google account icon.
6. Click `Manage Account`.
7. Go to `Security`.
8. Scroll to `Third-party apps with account access`.
9. Remove the previous authorization for this script.

### Testing the Integration

1. Go back to your spreadsheet.
2. Refresh the page.
3. Try adding a new task with a deadline.

## Feedback 

If you have any feedback, suggestions, or encounter any issues while using the platform, please don't hesitate to open an issue on GitHub. Your input is invaluable and helps us improve the platform for everyone.

<br/>
<p align="center">
  <a href="https://pulkitmathur.tech/"><img src="https://github.com/Pulkit1822/Pulkit1822/blob/main/animated-icons/pic.jpeg" alt="portfolio" width="32"></a>&nbsp;&nbsp;&nbsp;
  <a href="https://www.linkedin.com/in/pulkitkmathur/"><img src="https://github.com/TheDudeThatCode/TheDudeThatCode/blob/master/Assets/Linkedin.svg" alt="Linkedin Logo" width="32"></a>&nbsp;&nbsp;&nbsp;
  <a href="mailto:pulkitmathur.me@gmail.com"><img src="https://github.com/TheDudeThatCode/TheDudeThatCode/blob/master/Assets/Gmail.svg" alt="Gmail logo" height="32"></a>&nbsp;&nbsp;&nbsp;
  <a href="https://www.instagram.com/pulkitkumarmathur/"><img src="https://github.com/TheDudeThatCode/TheDudeThatCode/blob/master/Assets/Instagram.svg" alt="Instagram Logo" width="32"></a>&nbsp;&nbsp;&nbsp;
  <a href="https://in.pinterest.com/pulkitkumarmathur/"><img src="https://upload.wikimedia.org/wikipedia/commons/0/08/Pinterest-logo.png?20160129083321" alt="Pinterest Logo" width="32"></a>&nbsp;&nbsp;&nbsp;
  <a href="https://twitter.com/pulkitkmathur"><img src="https://upload.wikimedia.org/wikipedia/commons/5/57/X_logo_2023_%28white%29.png" alt="Twitter Logo" width="32"></a>&nbsp;&nbsp;&nbsp;
</p>

Happy learning and coding!

---

If you find this repository useful, don't forget to star it! ⭐️

### Written by [Pulkit](https://github.com/Pulkit1822)
