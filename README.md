# Defer Emails After Hours

An Outlook add-in that automatically delays emails sent outside of normal business hours.

## Features

- **After Hours Detection**: Automatically detects when emails are being sent outside of customizable business hours (default: 7 AM-6 PM, Monday-Friday)
- **Configurable Business Hours**: Allows users to set their own business hours through a settings taskpane
- **Automatic Delay**: Schedules emails to be sent at the start of the next business day
- **User Notification**: Keeps the user informed about the rescheduled delivery time

## Installation

1. Clone or download this repository
2. Install dependencies:
   ```
   npm install
   ```
3. Follow the [testing instructions](DEVELOPING.MD)

Learn more about sideloading an add-in into Outlook following [Microsoft's instructions](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing)


## How It Works

This add-in automatically manages emails sent after business hours:

1. **Automatic delay for after-hours emails**:
   - When composing a new email after business hours, the add-in automatically schedules it to be sent at the beginning of the next business day
     - For example, emails composed on Friday evening will be scheduled for Monday morning at your configured start time
     - US holidays are also recognized, ensuring emails aren't scheduled for delivery on non-business days
   - A persistent notification appears informing you when the email will be sent
   - No user action is required - this happens automatically for any email composed after hours
   
2. **Configurable business hours**:
   - Define your own business hours through the "Settings" button in the Outlook ribbon when in read mode
   - Select both start time (e.g., 7 AM, 8 AM, 9 AM) and end time (e.g., 5 PM, 6 PM, 7 PM)
   - Weekend days (Saturday and Sunday) are always considered outside business hours
   - US holidays are automatically recognized and handled as non-business days

3. **Manual delay removal**:
   - You can remove an automatic delay by clicking the "Remove Delay" button
   - This allows time-sensitive emails to be sent immediately when needed

### US Holiday Support

The add-in automatically recognizes the following US holidays as non-business days:

- New Year's Day (January 1)
- Martin Luther King Jr. Day (3rd Monday in January)
- Presidents Day (3rd Monday in February)
- Memorial Day (Last Monday in May)
- Independence Day (July 4)
- Labor Day (1st Monday in September)
- Thanksgiving Day (4th Thursday in November)
- Day after Thanksgiving (4th Friday in November)
- Christmas Eve (December 24)
- Christmas Day (December 25)

Emails composed during holidays or after hours will be scheduled for delivery on the next business day that is not a weekend or holiday.

## User Interface

- **Ribbon Integration**:
  - "Settings" button in read mode to access configuration
  - "Remove Delay" button in compose mode to cancel scheduled delays

- **Settings Panel**:
  - Enable/disable automatic delay feature
  - Configure business hours start and end times
  - Confirmation message of current settings

- **Notification System**:
  - Notifications when emails are delayed
  - Confirmation messages when settings are changed
  - Information about scheduled delivery times

## Development

### Prerequisites

- Node.js and npm
- Office Add-in development tools

### Project Structure

The add-in follows a modular architecture:

- **Shared Utilities** (`src/shared/utils.js`): Common code shared between components
  - Settings management (load/save)
  - Business hours calculation
  - Date/time formatting utilities
  - Common UI operations

- **Launch Event Handler** (`src/launchevent/launchevent.js`):
  - Handles new message compose events
  - Applies automatic delay when outside business hours
  - Interacts with the Outlook API for delay scheduling

- **Settings Taskpane** (`src/taskpane/taskpane.js`):
  - Provides user interface for configuring business hours
  - Enables/disables the automatic delay feature
  - Shows current configuration status

## License

[MIT License](LICENSE)
