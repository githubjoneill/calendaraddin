# Outlook Meeting Blank Fix (Outlook Agenda Enforcer)

This project implements an Outlook event-based add-in that leverages Microsoft's **Smart Alerts** feature (`OnAppointmentSend` event) to ensure that corporate meeting invitations are not sent without an agenda.

---

## Architectural Overview

When a user attempts to send a meeting invite (appointment), the Outlook runtime triggers the event-based add-in. The add-in runs a background JavaScript/TypeScript function to inspect the body of the meeting invite.

1. **Agenda Detection:** The JavaScript runtime asynchronously reads the meeting's body content.
2. **Evaluation:** It trims whitespace and checks if the body is empty, near-empty (e.g., fewer than 10 characters), or lacks standard agenda keywords.
3. **Smart Alert Enforcement:**
   - If a valid agenda/content is found, the send event is allowed to complete.
   - If the body is empty or invalid, the add-in automatically prepends a structured meeting agenda template (e.g., `Agenda:\n- [Topic 1]\n- [Topic 2]`) to the meeting body, cancels the send action, and displays an informative Smart Alert info bar to the user:
     > *"We've added a meeting agenda template to your invite body. Please complete the agenda before sending."*

---

## Technical Stack & Key Components

The add-in consists of the following key files:

1. **`manifest.xml`**: The application manifest defining the add-in identity, permissions, and registration of the `OnAppointmentSend` launch event.
2. **`src/launchevent/launchevent.js`**: Contains the execution logic for inspecting the appointment body, appending the agenda template if empty, and completing the event with the appropriate Smart Alerts parameters.
3. **`src/taskpane/taskpane.html` & `taskpane.js`**: Fallback user interface required by modern Office Add-ins (even if not actively used for background event processing, Office requires taskpane endpoints).
4. **`package.json`**: Node.js project configuration specifying dependencies (Office JS types, Webpack, etc.).
5. **`webpack.config.js`**: Bundles the code for production, ensuring `launchevent.js` is bundled into a self-contained script for the Office runtime environment.

---

## Project Structure

The codebase should be structured as follows:
```text
OutlookMeetingBlankFix/
├── GEMINI.md                    # This instructions & architectural guide
├── manifest.xml                 # Office Add-in manifest XML
├── package.json                 # Node dependencies and build scripts
├── webpack.config.js            # Webpack configuration for bundling JavaScript
├── src/
│   ├── launchevent/
│   │   └── launchevent.js       # Background event handler for OnAppointmentSend
│   └── taskpane/
│       ├── taskpane.html        # Fallback taskpane UI
│       └── taskpane.js          # Taskpane activation script
```

---

## Development & Testing Workflow

### 1. Prerequisites
- **Node.js** (v18 or newer recommended).
- **Microsoft 365 Developer Tenant** with administrative privileges.
- **Outlook Desktop (Windows/Mac)** or **Outlook on the Web** configured with your developer account.

### 2. Local Installation & Build
```powershell
# Install development dependencies
npm install

# Build the project (generates /dist folder)
npm run build

# Start the local development server (hosts manifest and code at https://localhost:3000)
npm run start
```

### 3. Sideloading for Testing
To test the add-in in your Microsoft Dev environment:
1. Go to [Outlook on the Web](https://outlook.office.com).
2. Open **Settings** (gear icon) > **View all Outlook settings** > **Mail** > **Customize actions**.
3. Under **Add-ins**, check the box for **Get Add-ins** (or open a new message, click the `...` menu, and select **Get Add-ins**).
4. Select **My add-ins** from the sidebar.
5. Under **Custom add-ins**, click **Add a custom add-in** > **Add from file...** and select your `manifest.xml`.
6. Open your calendar, create a new meeting, leave the body blank, and click **Send** to verify the Smart Alert prompts you and prepends the agenda template.

---

## Corporate Deployment Guide

To deploy this solution globally or to a specific subset of users in your corporate environment, an administrator must complete the following steps:

### Phase 1: Host the Production Build
Office Add-ins are web applications hosted on a web server.
1. Build the production assets using `npm run build`.
2. Upload the contents of the `dist/` directory (including `launchevent.js` and `taskpane.html`) to a secure, enterprise-grade, HTTPS-enabled web host (e.g., Azure App Service, Azure Static Web Apps, AWS S3 with CloudFront, or an internal IIS server).
3. Ensure the web server supports CORS and serves files over TLS 1.2 or higher.
4. Update the URLs in `manifest.xml` to point from `https://localhost:3000/` to your hosted production URLs (e.g., `https://addins.yourcompany.com/outlook-agenda-fix/`).

### Phase 2: Centralized Deployment via Microsoft 365 Admin Center
1. Navigate to the [Microsoft 365 Admin Center](https://admin.microsoft.com).
2. Go to **Settings** > **Integrated apps**.
3. Click **Add-ins** at the top right, or click **Deploy Add-in**.
4. Click **Next** on the prompt, then select **Upload custom apps**.
5. Choose **I have the manifest file (.xml) on my device**, upload your production `manifest.xml` file, and click **Next**.
6. **Assign Users:** Specify who will receive the add-in:
   - **Everyone**: Deploys to the entire organization.
   - **Specific users/groups**: Ideal for pilot testing (e.g., IT department, specific administrative units).
   - **Just me**: For final isolated validation.
7. **Deployment Method:** Select **Fixed (Default)** to automatically enable and install the add-in for all targeted users without allowing them to turn it off. This is highly recommended for compliance/standardization tools.
8. Click **Save** and wait for the deployment process to complete. It can take up to 24 hours for the add-in to propagate to all users' Outlook clients (both desktop and web).
