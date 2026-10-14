# Frame.io for PowerPoint and Word

A shared Office task-pane add-in for uploading PowerPoint presentations and Word documents to Frame.io. Sign in with your own Adobe account, choose an account, project, and folder, and upload the native file. A tracking ID saved in the document links later uploads to the same Frame.io version history.

The frontend uses Office.js and React. Adobe App Builder handles authentication and upload history; document bytes upload directly to Frame.io, without conversion or staging.

## Support

| Host | File | Required APIs | Preview |
| --- | --- | --- | --- |
| PowerPoint, existing compressed-export-capable hosts | `.pptx` | CompressedFile 1.1, PowerPointApi 1.3 | First slide when PowerPointApi 1.8 is available |
| Word on Mac and Windows desktop | `.docx` | CompressedFile 1.1, WordApi 1.3 | Deferred; document name only |
| Word web/mobile and Excel | Not supported | Upload disabled | None |

The Word path is implemented, but native acceptance on both desktop platforms remains a release gate. Automated/browser mocks do not establish Office export fidelity or live Frame.io DOCX acceptance. No minimum Office build is claimed beyond checked requirement sets. Legacy, macro-enabled, and template formats are rejected rather than relabeled as DOCX/PPTX.

## Requirements

- Node.js 22, npm, and OpenSSL for local HTTPS certificates.
- A supported PowerPoint or desktop Word installation with Office add-in sideloading enabled and document editing permitted when creating tracking metadata.
- Your own Adobe App Builder project with Runtime, State, Files, and web hosting available.
- Frame.io API v4 access and OAuth Web App credentials for user-based sign-in.

## Setup

After cloning the repository:

```sh
npm ci
npm install -g @adobe/aio-cli
aio plugins:install @adobe/aio-cli-plugin-app
cp .env.example .env
aio app use /path/to/your-console-export.json --merge
```

Download your project configuration from Adobe Developer Console and keep that credential-bearing export outside the repository. Import your own workspace, not another developer's configuration.

Complete the local environment file using the [environment template](.env.example):

- Set your Runtime namespace and management credential.
- Set `FRAMEIO_CLIENT_ID` and `FRAMEIO_CLIENT_SECRET` from your OAuth Web App credentials; retain the default scopes unless your integration requires different ones.
- Set `FRAMEIO_REDIRECT_URI` to the callback URL for the environment you will use.
- Generate `FRAMEIO_TOKEN_KEY` with `openssl rand -hex 32`. Keep it private and preserve it across restarts and deployments.

Never put credentials in frontend files or commit Console exports, private keys, or real environment files.

## Run Locally

Use `https://localhost:3000` both inside the container and on its host OS. Port 3000 is dedicated to this container: publish or forward container port 3000 to host port 3000, never 3002. No custom hostname aliases, DNS mappings, or hosts-file entries are needed.

OAuth currently runs on hosted Stage, whose origin and callback remain unchanged. Starting the local development server does not require switching OAuth to localhost. Only if explicitly testing local OAuth, register `https://localhost:3000/oauth-callback.html` in your OAuth credentials and set the local `FRAMEIO_REDIRECT_URI` to it. Local sign-in requires that opt-in configuration; a Stage callback does not authenticate the localhost panel.

```sh
npm run certs
npm run dev
```

Trust the generated public certificate on the computer running Office. Do not share the TLS private key. Open `https://localhost:3000/index.html` to check that the app loads without certificate errors. Existing certificate filenames are retained; an already trusted certificate covering `localhost` need not be regenerated.

Generate and validate the local Office manifest, then sideload it using your platform's Office add-in procedure:

```sh
npm run manifest:local
npm run validate:manifest
```

The same manifest supports PowerPoint and Word. Sideload it separately in each application; replace an older PowerPoint-only manifest and restart Office if Word does not list the add-in. Keep `ReadWriteDocument` permission. Open a DOCX in desktop Word, not Word on the web.

## Deploy Your Workspace

Use your workspace's static hosting origin from Adobe Developer Console. Register its `/oauth-callback.html` URL in your OAuth credentials and set that exact URL as the OAuth action's `FRAMEIO_REDIRECT_URI` input in [app.config.yaml](app.config.yaml). This repository pins that input to the Stage callback so a plain `aio app deploy` cannot accidentally deploy the localhost callback from `.env`. Local development still reads `.env`; keep local and hosted redirect registrations separate if you use both environments.

Before building, update [the public endpoint configuration](web-src/public-config.json) with your own same-origin hosted action URLs. Also update the explicitly pinned hosted origin in [the authentication client](web-src/auth.ts). The action package paths remain `/api/v1/web/msofficetoframe/status` and `/api/v1/web/msofficetoframe/oauth` unless you rename the package.

Release the additive DOCX-capable backend before the Word-enabled frontend. Missing `documentType` on existing requests/records still means PPTX; registry paths and presentation IDs are unchanged. New versions must match the linked document type and destination. Complete the native acceptance checks before releasing Word support.

```sh
npm test
npm run build
aio app deploy --no-build --no-publish --no-log-forwarding-update
```

Using the actual HTTPS `index.html` URL reported by deployment, generate and validate your hosted manifest:

```sh
npm run manifest -- --url "$TASKPANE_URL" --output office/manifest.deployed.xml
npx office-addin-manifest validate office/manifest.deployed.xml
```

Set `TASKPANE_URL` before running these commands. Replace the previously sideloaded manifest with the hosted one and restart Office. Hosted login sessions and upload history are separate from local development; local history is not automatically migrated.

## Using the Add-in

Sign in, select your destination, choose a valid `.pptx` or `.docx` filename, and choose **Upload to Frame.io**. Files must be between 1 byte and 1 GiB. Keep the pane open until the upload finishes. Once an upload is recorded, the add-in offers **Add a new version** and a link to the existing Frame.io asset; its destination stays fixed.

Tracking uses `MSOFFICE_TO_FRAME_UUID`: a PowerPoint presentation tag or a Word custom document property. Save the document when prompted to retain its UUID after closing. The add-in does not force a save; supported unsaved documents can upload. Rename, Save As, and copies intentionally retain the same ID and version history. A malformed, removed, changed, or inaccessible ID stops uploads rather than substituting a browser/filename identity. There is no distributed coauthoring lock or independent-copy reset.

If the linked Frame.io file or version stack has been deleted, the history warning remains visible. When the original project is still accessible, select an account, project, and folder to enable a replacement upload. The backend rechecks the missing asset and selected destination, clears the stale active mapping, and links the existing document UUID to the replacement at version 1. Private audit records are retained, but old confirmation retries cannot restore the retired mapping. Permission failures, moved assets, and unverified history do not enable this recovery.

Word uses an Office sign-in dialog, as PowerPoint does. Each host may have separate browser storage, so signing in to PowerPoint does not guarantee Word is signed in. No Word thumbnail/PDF conversion or Frame.io native DOCX review rendering is promised.

Cancellation does not delete a file already created in Frame.io. Use **Check upload status** for uncertain completion, inspect the destination before acknowledging another file creation, and use history retry/refresh when lookup fails. Signed storage URLs and Adobe tokens must never be logged or shared.

## Native Acceptance

Run these checks on both Word Mac and Windows with disposable documents and an explicitly approved Frame.io test destination. Record Office versions/update channels and results. They remain pending until executed on native installations.

1. Export saved and unsaved DOCX, including current edits, headers/footers, tables, images, fields, comments, and tracked changes. Include a large multi-slice file and a spaced/non-ASCII filename. Reopen the export in Word and check fidelity; ZIP signatures alone are insufficient.
2. Create a UUID once, inspect it in the exported document, save/reopen, rename, and Save As/copy. Verify identity/history persists. Test malformed metadata, denied writes, read-only/protected files, and identity changes while the pane is open; do not bypass protection.
3. Test Office-dialog login, cancel/close/failure, session restoration, logout/expiry, and the Mac dropdown click behavior. Verify account/project/folder browsing.
4. Upload and download a DOCX; compare its contents. Upload versions two and three and verify one shared version stack, filename and fixed destination. Native DOCX preview is not an acceptance requirement.
5. Test cancellation during export/transfer, offline/403/retry, pending confirmation, history refresh, moved/deleted assets, and registration failures. Verify cleanup and accurate warnings about retained remote files.
6. Check narrow/wide panes, long names, no empty Word slide frame, and disabled unsupported hosts. Regress PowerPoint preview, OAuth, existing records without type, and first/second/third uploads against the new backend.