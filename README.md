# Frame.io for PowerPoint

A PowerPoint task-pane add-in for uploading presentations to Frame.io. Sign in with your own Adobe account, choose an account, project, and folder, and upload the presentation. A tracking ID saved in the presentation links later uploads to the same Frame.io version history.

The frontend uses Office.js and React. Adobe App Builder handles authentication and upload history; presentation bytes upload directly to Frame.io.

## Requirements

- Node.js 22, npm, and OpenSSL for local HTTPS certificates.
- A supported PowerPoint installation with Office add-in sideloading enabled.
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

Register `https://www.localappbuilder.com:3000/oauth-callback.html` in your OAuth credentials and use it as the local `FRAMEIO_REDIRECT_URI`. Configure `www.localappbuilder.com` to resolve to your development server; for a container, use the address reachable from the PowerPoint host.

```sh
npm run certs
npm run dev
```

Trust the generated public certificate on the computer running PowerPoint. Do not share the TLS private key. Open `https://www.localappbuilder.com:3000/index.html` to check that the app loads without certificate errors.

Generate and validate the local Office manifest, then sideload it using your platform's Office add-in procedure:

```sh
npm run manifest:local
npm run validate:manifest
```

## Deploy Your Workspace

Use your workspace's static hosting origin from Adobe Developer Console. Register its `/oauth-callback.html` URL in your OAuth credentials and use that exact URL as `FRAMEIO_REDIRECT_URI` for both the build and deployment. Keep local and hosted redirect registrations separate if you use both environments.

Before building, update [the public endpoint configuration](web-src/public-config.json) with your own same-origin hosted action URLs. Also update the explicitly pinned hosted origin in [the authentication client](web-src/auth.ts). The action package paths remain `/api/v1/web/msofficetoframe/status` and `/api/v1/web/msofficetoframe/oauth` unless you rename the package.

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

Set `TASKPANE_URL` before running these commands. Replace the previously sideloaded manifest with the hosted one and restart PowerPoint. Hosted login sessions and upload history are separate from local development; local history is not automatically migrated.

## Using the Add-in

Sign in, select your destination, and choose **Upload to Frame.io**. Save the presentation to retain its tracking ID. Once an upload is recorded, the add-in offers **Add a new version** and a link to the existing Frame.io asset. Keep the pane open until the upload finishes. Copies of a presentation inherit its tracking ID and version history.