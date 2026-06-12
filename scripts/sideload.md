<!-- SPDX-License-Identifier: Apache-2.0 -->

# Sideloading spark-connect-excel into Excel

Sideloading installs the add-in from `manifest.xml` without publishing to
AppSource. This is the standard development and team-deployment path.

## Recommended: Excel on the web (fewest steps, no admin)

1. Start the HTTPS dev server: `npm run dev:https` (run `npx office-addin-dev-certs
   install` once first).
2. Open **Excel on the web** in Microsoft Edge or Google Chrome.
3. **Insert -> Add-ins -> Upload My Add-in** -> choose `manifest.xml`.

That's it - no catalog, no admin, no desktop config. To share with others, host
the bundle and point the manifest at it (see `docs/distribution.md`), then they
upload that `manifest.xml` the same way.

## Quick start for desktop (Windows / Mac)

```bash
# Starts the dev server AND sideloads the add-in into desktop Excel:
npx office-addin-debugging start manifest.xml
```

This command:
1. Starts the Vite dev server (or uses an existing one)
2. Copies the manifest to the platform-appropriate location
3. Opens Excel with the add-in sideloaded

## Platform-specific manual steps

### Windows

1. Open Excel.
2. Go to **File > Options > Trust Center > Trust Center Settings > Trusted Add-in Catalogs**.
3. Add the URL of your manifest (e.g. `https://localhost:3000/manifest.xml`) as a trusted catalog.
4. Check **Show in Menu** -> OK.
5. Go to **Insert > My Add-ins > Shared Folder** and click **Spark Connect for Excel**.

Alternatively, copy `manifest.xml` to:
```
%APPDATA%\Microsoft\Excel\XLSTART\
```
Then restart Excel.

### Mac

```bash
# office-addin-debugging handles the Mac-specific manifest location:
npx office-addin-debugging start manifest.xml
```

Manual path:
```
~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/
```
Create the `wef/` directory if it doesn't exist, copy `manifest.xml` there, and restart Excel.

### Excel on the web (Microsoft 365)

1. Open Excel on the web.
2. Go to **Insert > Add-ins > Upload My Add-in**.
3. Browse to and upload `manifest.xml`.

The add-in's `SourceLocation` must be accessible from the internet (or via a
tunnel - see the ngrok tip below).

## HTTPS for Windows/Mac

The dev server uses plain HTTP by default. Windows WebView2 and Mac WKWebView
require HTTPS for the task pane. To enable HTTPS in dev:

```bash
npx office-addin-dev-certs install   # once - installs an OS-trusted local cert
npm run dev:https                    # serves https://localhost:3000 (COI headers + TLS)
```

## Testing with ngrok (Excel on the web / remote team)

If you need to sideload from a machine that isn't localhost:

```bash
npm run dev       # Start Vite on port 3000
ngrok http 3000   # Expose via HTTPS tunnel
```

Update `manifest.xml`'s `SourceLocation` to the ngrok HTTPS URL, then sideload
the updated manifest in Excel on the web.

## Stopping / uninstalling

```bash
npx office-addin-debugging stop manifest.xml
```

To remove the sideloaded add-in manually:
- Windows: delete the manifest file from `%APPDATA%\Microsoft\Excel\XLSTART\`
- Mac: delete the manifest from `~/Library/Containers/.../wef/`
- Excel on the web: Insert > My Add-ins -> context menu -> Remove
