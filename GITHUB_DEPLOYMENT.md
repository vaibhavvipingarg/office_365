# GitHub Pages Deployment Guide

Follow these steps to deploy the Office 365 Salesforce Add-in to GitHub Pages:

## 1. Push to GitHub

```bash
# Create a new repo on GitHub named "office_365" if you haven't already
# Then push your gh-pages branch

# Add the files
git add gh-pages-root manifest-gh-pages.xml
git commit -m "Add GitHub Pages deployment files"
git push origin gh-pages
```

## 2. Enable GitHub Pages

1. Go to your GitHub repository (https://github.com/vaibhavvipingarg/office_365)
2. Click on "Settings"
3. In the left sidebar, click on "Pages"
4. Under "Source", select the "gh-pages" branch and "/ (root)" folder
5. Click "Save"
6. Wait a few minutes for GitHub Pages to build your site

## 3. Update Salesforce Connected App

1. Log into Salesforce
2. Go to Setup > App Manager
3. Find your Connected App and edit it
4. Under "OAuth Settings", add a new Callback URL:
   ```
   https://vaibhavvipingarg.github.io/office_365/oauth/callback/
   ```
5. Save the changes

## 4. Test the Add-in

1. Download the manifest-gh-pages.xml file
2. In Word Online, go to Insert > Add-ins
3. Choose "Upload My Add-in"
4. Select your manifest-gh-pages.xml file
5. The add-in should now load from GitHub Pages

## Troubleshooting

If the add-in doesn't work:

1. Check browser console for errors (F12)
2. Verify GitHub Pages is properly configured
3. Ensure your Salesforce Connected App settings are correct
4. Confirm all URLs in the manifest file are correct 