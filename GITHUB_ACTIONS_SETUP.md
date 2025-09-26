# GitHub Actions Setup for Certificate Generator

## Overview
This guide will help you migrate your certificate generator from Render to GitHub Actions, which provides free SMTP functionality and is perfect for your use case.

## Prerequisites
- GitHub repository with your certificate generator code
- Google Cloud Console project with Sheets and Drive APIs enabled
- Gmail account for SMTP

## Step 1: Repository Setup

1. **Push your code to GitHub** (if not already done):
   ```bash
   git add .
   git commit -m "Add GitHub Actions support"
   git push origin main
   ```

2. **Ensure you have these files in your repository**:
   - `.github/workflows/certificate-generator.yml` ✅ (created)
   - `generate-certificates.js` ✅ (created)
   - `package.json` ✅ (existing)
   - All your existing files

## Step 2: Configure GitHub Secrets

Go to your GitHub repository → Settings → Secrets and variables → Actions, and add these secrets:

### Required Secrets:
```
CLIENT_ID=your_google_client_id
CLIENT_SECRET=your_google_client_secret
REDIRECT_URI=your_redirect_uri
REFRESH_TOKEN=your_refresh_token
TEMPLATE_ID=your_google_slides_template_id
FOLDER_ID=your_google_drive_folder_id
EMAIL=your_gmail_address
PASSWORD=your_gmail_app_password
```

### How to get these values:

1. **Google API Credentials** (if you don't have them):
   - Go to [Google Cloud Console](https://console.cloud.google.com/)
   - Create a new project or select existing one
   - Enable Google Sheets API and Google Drive API
   - Go to "Credentials" → "Create Credentials" → "OAuth 2.0 Client ID"
   - Set application type to "Web application"
   - Add your redirect URI (can be `http://localhost:3000` for testing)
   - Download the JSON file and extract `client_id` and `client_secret`

2. **Refresh Token** (if you don't have one):
   - Use the OAuth playground: https://developers.google.com/oauthplayground/
   - Select Google Sheets API and Google Drive API
   - Authorize and get the refresh token

3. **Gmail App Password**:
   - Go to your Google Account settings
   - Security → 2-Step Verification (enable if not already)
   - App passwords → Generate app password for "Mail"
   - Use this password (not your regular Gmail password)

## Step 3: Usage

### Method 1: Manual Trigger (Recommended)
1. Go to your GitHub repository
2. Click on "Actions" tab
3. Select "Certificate Generator" workflow
4. Click "Run workflow"
5. Fill in the required parameters:
   - **Sheet ID**: Your Google Sheet ID (from the URL)
   - **Sheet Name**: Name of the sheet (e.g., "Sheet1")
   - **Start Date**: Format YYYY-MM-DD (e.g., "2024-01-15")
   - **End Date**: Format YYYY-MM-DD (e.g., "2024-01-20")
   - **School**: Optional, leave empty for all schools
6. Click "Run workflow"

### Method 2: Webhook Trigger (Advanced)
If you want to trigger from an external system, you can modify the workflow to accept webhook triggers.

## Step 4: Monitor Execution

1. **View Progress**: Click on the running workflow to see real-time logs
2. **Download Results**: After completion, go to the "Artifacts" section to download:
   - `certificates.pdf` (if no school specified)
   - `certificates.zip` (if school specified)
3. **Check Logs**: View detailed logs for any errors

## Step 5: Email Delivery

- **Individual Emails**: When no school is specified, certificates are sent individually via email
- **Bulk Download**: When a school is specified, certificates are packaged in a ZIP file for download

## Troubleshooting

### Common Issues:

1. **"Invalid credentials" error**:
   - Check that all secrets are correctly set
   - Verify Gmail app password (not regular password)
   - Ensure Google APIs are enabled

2. **"Sheet not found" error**:
   - Verify the Sheet ID is correct
   - Ensure the sheet is shared with your Google account
   - Check the sheet name matches exactly

3. **"Email sending failed"**:
   - Verify Gmail app password
   - Check that 2FA is enabled on Gmail
   - Ensure the email addresses in the sheet are valid

4. **"Template not found"**:
   - Verify TEMPLATE_ID is correct
   - Ensure the Google Slides template is shared with your account

### Debug Steps:
1. Check the workflow logs for detailed error messages
2. Test with a small dataset first
3. Verify all environment variables are set correctly
4. Test Google API access independently

## Cost Analysis

### GitHub Actions (Recommended):
- **Cost**: FREE (2,000 minutes/month)
- **Your usage**: ~5-10 minutes per certificate batch
- **Monthly capacity**: 200-400 certificate batches
- **Maintenance**: None required

### AWS Alternative:
- **Cost**: $5-20/month minimum
- **Complexity**: High (EC2, load balancer, domain, SSL)
- **Maintenance**: Regular updates and monitoring required

## Migration Benefits

✅ **Free SMTP functionality**  
✅ **No server management**  
✅ **Automatic scaling**  
✅ **Reliable infrastructure**  
✅ **Easy to use**  
✅ **Built-in logging and monitoring**  
✅ **Secure secret management**  

## Next Steps

1. Set up the secrets in your GitHub repository
2. Test with a small batch of certificates
3. Update your frontend to trigger the workflow (optional)
4. Monitor the first few runs to ensure everything works correctly

## Support

If you encounter any issues:
1. Check the GitHub Actions logs first
2. Verify all secrets are correctly configured
3. Test with a minimal dataset
4. Review the troubleshooting section above

The GitHub Actions approach is much simpler and more cost-effective than AWS for your certificate generation needs!
