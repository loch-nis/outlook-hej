# Hej-hilsen

Outlook add-in that writes `Hej <fornavn>` at the top of a mail you compose, based on the To field.
Works in Outlook on the web, new Outlook for Mac, and new and classic Outlook for Windows (Microsoft 365).

## What it writes

| To field | Greeting |
| --- | --- |
| Anne Hansen | Hej Anne |
| Anne Hansen, Peter Holm | Hej Anne og Peter |
| Anne, Peter, Mette | Hej Anne, Peter og Mette |
| 4 or more, or a distribution list | Hej alle |
| info@firma.dk | Hej |

- First name comes from the display name ("Hansen, Anne" and "ANNE HANSEN" both give Anne), or from `anne.hansen@` style addresses when there is no display name.
- Cc and Bcc are ignored.
- The greeting updates as you add or remove To recipients, until you edit it. After that the add-in leaves it alone.
- If the mail already starts with a greeting you wrote (Hej, Kære, Hi, Dear, ...), nothing is inserted.
- Replies and forwards get the greeting above the quoted mail.

Change the wording in `CONFIG` at the top of [src/launchevent.js](src/launchevent.js), for example `punctuation: ","` for "Hej Anne,".

## Build and host

The add-in is static files served over HTTPS from GitHub Pages. Every push to `main` runs the tests and deploys (see [.github/workflows/pages.yml](.github/workflows/pages.yml)):

- Runtime page: https://loch-nis.github.io/outlook-hej/commands.html
- Manifest to install: https://loch-nis.github.io/outlook-hej/manifest.xml

Locally:

```bash
npm test
node scripts/build.js https://loch-nis.github.io/outlook-hej
```

This writes `docs/` and `manifest.xml` (both git-ignored; CI builds its own).

## Install for yourself

1. Open https://aka.ms/olksideload (Outlook on the web opens the add-ins dialog).
2. My add-ins → Custom add-ins → Add a custom add-in → Add from file → pick `manifest.xml` (download it from the link above).
3. Reload Outlook. The add-in follows your mailbox, so new Outlook for Mac and Windows pick it up too (classic Windows can take up to 24 hours).

On Mac 16.85 and later, Get Add-ins opens the Marketplace instead, so install through Outlook on the web.

## Install for the whole company

Microsoft 365 admin center → Settings → Integrated apps → Upload custom apps → App type **Office Add-in** → upload `manifest.xml` → assign users → deploy as **Fixed**. It can take up to 24 hours to reach everyone.

## Updating

Changes to `src/` only: bump `version` in `package.json` and push. Outlook picks up the new script within about 10 minutes (GitHub Pages cache), or on restart for classic Windows. No reinstall.

Changes to the manifest template: push, then remove and re-add the add-in (or Update it in the admin center, which asks for consent again).

## How it edits the body

- The first greeting goes in with `prependAsync`, which leaves the rest of the body alone.
- Updating it (Anne → Anne og Peter) means rewriting the body with `setAsync`. That is skipped when the body holds attached inline images, because Outlook can drop them ([office-js#6808](https://github.com/OfficeDev/office-js/issues/6808), [#6944](https://github.com/OfficeDev/office-js/issues/6944)). Linked (https) images are fine.
- Recipient changes are handled 1 second after the last one, so the greeting doesn't change on every keystroke in the To field.
- The add-in finds its greeting by text and remembers it in `sessionData`. It does not tag it with an id or class, because Outlook rewrites those ([office-js#5296](https://github.com/OfficeDev/office-js/issues/5296)).

## Limits

- On Mac the cursor disappears after the greeting is inserted; click in the body to continue (documented Outlook behaviour).
- Closing an untouched reply asks whether to save it, because the greeting counts as an edit.
- Reopening a saved draft doesn't fire the compose event; changing its recipients still does.
- Double first names ("Anne Marie Hansen") give the first word only.
- Outlook mobile isn't covered.
- Classic Outlook for Windows needs version 2403 or later (the script uses modern JavaScript).
- Known Outlook bugs that can stop the events: [office-js#6956](https://github.com/OfficeDev/office-js/issues/6956) (compose event not firing on web and Mac 16.113, one report), [#5967](https://github.com/OfficeDev/office-js/issues/5967) (Mac, compose inside the main window), [#3861](https://github.com/OfficeDev/office-js/issues/3861) (web, popping the compose window out quickly).
