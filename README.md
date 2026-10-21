# New to All (Outlook add-in)

When you are **reading** an email, this add-in adds a **New to All** button on the message ribbon. Clicking it opens a **new compose** window (not a reply or forward) with:

- **To:** the original sender plus everyone on the original **To** line  
- **Cc:** everyone on the original **Cc** line  
- Your own address is removed; addresses are deduplicated without regard to case  
- Subject and body are left empty (no quoted thread)

Works in **new Outlook for Windows** and **Outlook on the web**.

## Host the add-in (GitHub Pages)

Static files live in [`/docs`](docs/). After you push to `main`:

1. Open the repo on GitHub → **Settings** → **Pages**.
2. Under **Build and deployment**, set **Source** to **Deploy from a branch**.
3. Choose branch **main** and folder **/docs**, then save.

The add-in is served from:

`https://havedill.github.io/outlook-new-to-all/`

(Enable Pages once; until then, sideloading will fail to load the function file.)

## Sideload in new Outlook for Windows

1. Download [`manifest.xml`](https://havedill.github.io/outlook-new-to-all/manifest.xml) (or use the copy in the repo root).
2. In Outlook, open **Apps** or **Get Add-ins** (exact label can vary by build).
3. Go to **My add-ins** → **Add a custom add-in** → **Add from file** and select `manifest.xml`.
4. Open a message in **read** view; you should see **New to All** on the ribbon (often under the default message tab).

## Pin the button

If the command does not stay visible, use **Customize toolbar** / **Pin** on the ribbon (wording varies in new Outlook). Pin **New to All** so it appears while reading messages.

## Development

- [`manifest.xml`](manifest.xml) — sideload manifest (duplicate in [`docs/manifest.xml`](docs/manifest.xml) for download via Pages).
- [`docs/function-file.html`](docs/function-file.html) — function command host page.
- [`docs/commands.js`](docs/commands.js) — `displayNewMessageForm` logic.

Validate the manifest:

```bash
npx office-addin-manifest validate manifest.xml
```

Regenerate PNG icons (optional):

```bash
node scripts/generate-icons.mjs
```
