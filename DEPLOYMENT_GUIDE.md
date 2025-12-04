# 🚀 Apps Script Web App Deployment Guide

## ⚠️ CRITICAL: You MUST use a Production Deployment (/exec URL)

The "Google Docs encountered an error" you're experiencing is caused by using a **development URL** (`/dev`) instead of a **production deployment** (`/exec`).

---

## 📋 Step-by-Step Deployment Instructions

### Step 1: Push Your Code to Apps Script

Since you're using clasp, push your latest changes:

```bash
cd /home/user/HR_APPSCRIPT_CLAUDE
clasp push
```

### Step 2: Open Apps Script Editor

```bash
clasp open
```

Or go to: https://script.google.com

### Step 3: Authorize Required Permissions

Before deploying, you need to grant all OAuth permissions:

1. In the Apps Script Editor, select **`forceAuthorization`** from the function dropdown
2. Click **Run** (▶️ button)
3. Click **Review Permissions**
4. Select your Google account
5. Click **Advanced** → **Go to [Your Project] (unsafe)**
6. Review and **Allow** all permissions:
   - ✅ View and manage spreadsheets
   - ✅ View and manage documents
   - ✅ View and manage Drive files
   - ✅ Connect to external services

### Step 4: Create a NEW Production Deployment

**IMPORTANT:** Don't use "Test deployments" - they use `/dev` URLs and cause errors.

1. Click **Deploy** → **New deployment** (top right)
2. Click the gear icon ⚙️ next to "Select type"
3. Choose **Web app**
4. Configure settings:
   - **Description:** `HR Payroll System - Production v1`
   - **Execute as:** `Me (your-email@gmail.com)`
   - **Who has access:** `Anyone` (or `Anyone with Google account` if you prefer)
5. Click **Deploy**
6. **Authorize** if prompted
7. **Copy the Web app URL** - it should end with `/exec`

### Step 5: Verify the URL

Your production URL should look like:
```
https://script.google.com/macros/s/AKfycbx.../exec
                                              ^^^^
                                              Must be /exec NOT /dev
```

**❌ WRONG (Development URL):**
```
https://script.google.com/macros/s/AKfycbx.../dev
```

**✅ CORRECT (Production URL):**
```
https://script.google.com/macros/s/AKfycbx.../exec
```

### Step 6: Test the Deployment

1. **Close all browser tabs** with the old URL
2. Open an **incognito/private window**
3. Navigate to your **new `/exec` URL**
4. You should see the login page (no Google Docs error!)
5. Try logging in with your credentials

---

## 🔧 Troubleshooting

### If you still see "Google Docs encountered an error":

1. **Verify you're using the `/exec` URL** (not `/dev`)
2. **Clear browser cache** or use incognito mode
3. **Check Apps Script Executions log:**
   - Open Apps Script Editor
   - Click **Executions** (clock icon in left sidebar)
   - Look for errors in the latest execution
   - Share the error message if you find one

### If you see a blank screen:

1. Open browser **Developer Tools** (F12)
2. Go to the **Console** tab
3. Look for JavaScript errors
4. Check the **Network** tab for failed requests

### If login fails:

Run these diagnostic functions in Apps Script Editor:

```javascript
// Check if users exist
listAllUsers()

// Test authentication
testLoginFlow()

// Verify deployment
checkDeploymentStatus()
```

---

## 📝 Quick Reference

### Managing Deployments

- **View all deployments:** Deploy → Manage deployments
- **Update deployment:** Click Edit (pencil icon) → New version → Deploy
- **Get deployment URL:** Manage deployments → Copy URL

### Common Issues

| Issue | Cause | Solution |
|-------|-------|----------|
| "Google Docs error" | Using `/dev` URL | Create new deployment with `/exec` URL |
| Blank screen | Browser cache | Clear cache or use incognito |
| Login fails | No users | Run `createTestUser()` function |
| Permission errors | Missing OAuth scopes | Run `forceAuthorization()` |

---

## ✅ Verification Checklist

Before considering deployment complete:

- [ ] Ran `clasp push` to upload latest code
- [ ] Ran `forceAuthorization()` and granted all permissions
- [ ] Created a **NEW deployment** (not test deployment)
- [ ] Verified URL ends with `/exec` (not `/dev`)
- [ ] Tested in incognito/private window
- [ ] Successfully logged in and saw dashboard
- [ ] No "Google Docs" errors appear

---

## 🆘 Still Having Issues?

If you've followed all steps and still have problems:

1. Share the **exact URL** you're using (with sensitive parts redacted)
2. Share any **error messages** from:
   - Browser console (F12 → Console tab)
   - Apps Script Executions log
3. Confirm which step in this guide is failing

---

**Last Updated:** 2025-12-04
**Version:** 2.0.4-production-deployment-required
