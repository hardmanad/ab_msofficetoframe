# Authentication, Secrets, and Token Storage

This describes the implemented OAuth foundation as of October 13, 2026. No credential values are included.

"Connected" means the backend has completed OAuth and stored a session for the browser. Successfully loading choices confirms read access, but does not prove upload permission. Direct upload now creates a file using the signed-in user's token; Frame.io enforces write permission for that request. No file bytes or provider tokens pass through the browser-to-Runtime upload-control requests.

At the time of inspection, there was one encrypted local session record. Both that record and [.env](.env) had owner-only `600` permissions. The number of records changes as users sign in and out.

## Credential Inventory

| Value | Purpose and storage |
| --- | --- |
| `FRAMEIO_CLIENT_ID` | Identifies the Adobe OAuth Web App. Stored in `.env`, copied from the Console export. Not a secret; it appears in the authorization URL. |
| `FRAMEIO_CLIENT_SECRET` | Proves the backend represents that Web App during code exchange, refresh, and revocation. Stored in `.env`; never sent to the panel. |
| `FRAMEIO_TOKEN_KEY` | A separately generated, random 256-bit encryption key. Stored in `.env` and preserved across configuration runs. Encrypts backend token and login records. |
| `AIO_runtime_auth` | Adobe Runtime management credential used by CLI/build/deployment tooling. Not a user-login token and never sent to the browser. |
| `IMS_OAUTH_S2S_*` | Credentials for the separate service account imported from the export. Not used by panel login. |
| `AIO_ims_contexts_*` | CLI-generated credential copies and context settings. The `FRAMEIO_*` aliases explicitly select the Web App credential. |
| `FRAMEIO_REDIRECT_URI`, `FRAMEIO_SCOPES` | Non-secret configuration: where Adobe returns the browser and which permissions the app requests. |

The original [Console export](msofficetoframe-27200-Stage.json) also contains plaintext credentials. It is ignored by Git, but downloaded copies and backups must be protected. The inspected [.aio](.aio) contains project metadata and credential references, not secret values.

The private configuration script, [scripts/configure-oauth.cjs](scripts/configure-oauth.cjs), selects the Web App credential from the export, preserves unrelated environment values, and generates the encryption key only if it is absent. Client IDs are identifiers, not passwords; client secrets, Runtime auth, and the encryption key are confidential.

## Browser Session Key

The panel generates a random 256-bit key and stores it in browser `localStorage` under `frameio.app-session`. A valid key from the previous tab-scoped `sessionStorage` implementation is migrated without changing it. The pane reuses the key for its automatic backend session check when reopened; Adobe tokens remain server-side.

This is an app-session bearer credential, not an Adobe token. The panel sends it to the backend as `Authorization: Bearer ...`. The backend hashes it with SHA-256 to identify the corresponding token record.

Someone who steals this key could use that app session, so it is sensitive. No Adobe access token, refresh token, client secret, or encryption key is stored in browser storage. Persistent storage is scoped to the add-in origin and browser/Office webview profile, not the presentation. It is shared by panels using that same storage partition and can survive panel/application closure. Office storage policies, private browsing, cache clearing, a different profile, or a changed origin may prevent restoration. If persistent storage is blocked, the pane falls back to tab-scoped session storage; if both are blocked, sign-in fails explicitly rather than issuing changing handles.

Persistence increases the exposure window to XSS and other users of the same device/profile. Use explicit Sign out on shared machines. This browser handle does not extend the backend's existing 24-hour session lifetime or make expired/revoked Adobe refresh tokens valid.

The browser-side implementation is in [web-src/auth.ts](web-src/auth.ts).

Word and PowerPoint both use Office dialogs for sign-in, independently of whether native file export is available. Plain browser login retains its popup flow. Word and PowerPoint may use separate Office storage partitions: a connection in one host does not guarantee a restored connection in the other. Verify session restoration in each native host.

## OAuth Tokens

During login:

1. The backend creates a ten-minute transaction containing a PKCE verifier, nonce, and session association. It stores that record encrypted. The authorization request carries an opaque state value, nonce, and S256 PKCE challenge, not the verifier or client secret.
2. Adobe returns a short-lived authorization code through the callback URL. The callback sends it to the backend and removes the URL query before loading Office.js.
3. The backend exchanges the code using the Web App secret and PKCE verifier, validates the ID token's issuer, audience, and nonce over the authenticated token-endpoint response, then stores the access token and refresh token encrypted. The ID token is not retained.

The access token authorizes API calls as the signed-in user, subject to that user's Frame.io permissions and the granted OAuth scopes. The refresh token lets the backend obtain a new access token without repeating interactive login. Neither token is returned to the panel.

Login transactions are removed before exchange. Concurrent callback replay is additionally constrained by Adobe's single-use code and PKCE enforcement; the Runtime State implementation does not provide atomic compare-and-delete here.

Currently, encrypted records live in [.local-state/](.local-state/). They use AES-256-GCM with a random initialization vector, authentication tag, and binding to the record's identity. The outer file contains ciphertext and expiration metadata, not plaintext tokens. The action implementation is in [actions/oauth/index.js](actions/oauth/index.js); the local storage adapter is in [scripts/local-state.cjs](scripts/local-state.cjs).

Stage was deployed on 2026-10-13. Hosted encrypted records use Adobe Runtime State. Credential and encryption settings are private Runtime action inputs, configured through [app.config.yaml](app.config.yaml). Runtime State uses the backend's native Runtime credentials, not the browser's app-session key. Hosted OAuth bootstrap was verified; full provider sign-in remains pending the registered hosted redirect. Local sessions are not migrated.

## Expiry and Sign-Out

- Access-token expiry follows Adobe's response. A session check refreshes it when fewer than 60 seconds remain. Refresh is request-driven, not a background timer.
- Backend session records expire 24 hours after the latest successful exchange or refresh.
- Upload records bind the upload ID, destination, file ID and signed part URLs to the app session. They are encrypted using the same authenticated-record storage and expire 24 hours after creation. Status requests require a live session. This retention does not extend the provider URLs' lifetime; logout invalidates access through the deleted session, but does not immediately erase upload records or revoke signed storage URLs.
- The durable upload journal is separate from expiring encrypted token/upload state. Runtime uses private App Builder Files; local development uses owner-only JSON files under `.local-state/registry/`. Journal/current-index records have no TTL and contain metadata and Frame.io IDs/links, not tokens or signed upload URLs. Logout does not remove history. Lookup requires a live login and verifies current Frame.io access before returning the record. Registry storage is private, but its metadata is not AES-encrypted with `FRAMEIO_TOKEN_KEY` by this application; native storage protection and filesystem permissions apply. Stage is deployed; local-to-cloud history migration has not been performed.
- Closing the pane or PowerPoint does not deliberately remove the persistent browser handle, revoke Adobe tokens, or delete backend records. Reopening checks the backend session before showing Connected and refreshes provider access when necessary. Actual persistence across a native PowerPoint restart requires verification in the user's Office profile.
- Sign-out deletes the backend session, attempts Adobe revocation, and clears both persistent and legacy tab-scoped app-session keys. Failed revocation is reported explicitly rather than represented as confirmed provider revocation.
- Expired local files are deleted when accessed, rather than by a background cleanup job. Runtime State uses storage TTLs.
- The encryption key must remain stable across server restarts and deployments. Changing it makes existing encrypted records unreadable; rotation needs a deliberate session invalidation and cleanup procedure.

## HTTPS Certificate

[.certs/localappbuilder.key](.certs/localappbuilder.key) is the server's TLS private key. It protects HTTPS transport, not OAuth token storage. Keep it private and do not transfer it merely to establish browser trust.

The corresponding [public certificate](.certs/localappbuilder.crt) is what is imported and trusted on the Mac. New certificates cover `localhost` and `127.0.0.1`; existing trusted certificates covering `localhost` can be reused. The legacy filenames do not imply hostname aliases. Local development uses only `https://localhost:3000` inside the container and on the host. The generated certificate is self-signed with a 30-day validity; regeneration requires trusting the new certificate on the host.

The hosts-file entry only controls hostname resolution. It does not establish certificate trust or ownership of the public domain. Use an owned domain and appropriate certificates for production.

## Security Boundaries

- Adobe account passwords are entered on Adobe's login pages, not collected or stored by this app.
- The app does not rely on an Experience Cloud shell, `sharedContext`, or an automatically available IMS token. Adobe IMS is used for the explicit user-login flow.
- The OAuth action accepts only start, callback, session, logout, accounts, projects, folders, presentation-lookup, upload-start and upload-status operations from the configured origin. Browsing, lookup and versioning require the app-session bearer credential and use the signed-in user's token, not a service-account credential.
- Browsing uses fixed GET routes on `https://api.frame.io/v4`. IDs are validated, folder/project ownership is checked, and pagination links are never followed as arbitrary URLs. Only normalized choice IDs/names and navigation metadata are sent to the panel; no provider tokens or raw error details are returned. Selections remain in frontend memory and are reset when their parent selection changes or the user signs out.
- Upload creation uses a fixed Frame.io POST route after destination checks. Only validated HTTPS S3 part URLs and sizes are returned. These URLs are temporary write capabilities: do not log, persist in browser storage, or share them. PUTs carry no app/provider bearer credential or cookies. The frontend keeps the active upload ID in memory; no resume after pane closure is promised. Cancel does not delete partial Frame.io files or revoke URLs.
- Document identity is a generated UUID in `MSOFFICE_TO_FRAME_UUID`, using PowerPointApi 1.3 presentation tags or WordApi 1.3 custom document properties. Creation requires document-write permission and a user save for persistence after closing. It is an identifier, not an authorization credential. Renamed/copied documents retain the same UUID and history; possessing it does not bypass Frame.io permission checks. Identity is checked again before export. Confirmed version registration uses fixed version-stack create/file-move routes, with the existing destination verified again before upload. DOCX/PPTX types are checked at upload-start and finalization; missing legacy type means PPTX, regardless of remote display name. Publicly returned view links are restricted to HTTPS Frame.io hosts.
- Web-action `final: true` protects configured credential inputs from request overrides.
- The local server parses `.env` for the backend only; it does not inject those values into Parcel. Frontend environment inlining is disabled, and builds scan generated web assets for known secret values.
- The secret scanner excludes the Console export's public `redirect_uri` and `defaultRedirectUri` metadata from secret collection; credentials and tokens remain scanned.
- Environment files (except the placeholder `.env.example`), Adobe/AWS credentials and backups, Console exports matching `*-[0-9]*-*.json` and their backups, local token/upload history, private-key files, local presentations, editor backups, and generated artifacts are excluded by [.gitignore](.gitignore). The export pattern follows the project-name, numeric-org-ID, and environment naming convention; it can also match unrelated JSON files with that shape. Console exports renamed outside that filename pattern still require explicit ignore rules. Ignore rules do not remove files already tracked by Git; review staged content before the first push.

Encryption does not protect against complete container compromise: someone who can read both `.env` and the encrypted records can decrypt the tokens. Git-ignore rules prevent accidental commits; they are not an access-control boundary. Avoid logging authorization headers, tokens, codes, or callback query strings.

Before external rollout, add infrastructure rate limits for public OAuth bootstrap, test concurrent refresh/logout behavior, verify deployed CORS and State access, and establish key rotation, revocation, and audit procedures. The integration foundation is not a production security certification.

See [README.md](README.md) for setup, certificate trust, destination selection, upload limits and recovery behavior, manifest generation, deployment boundaries, and remaining manual verification.