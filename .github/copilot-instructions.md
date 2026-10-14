# Local development

- Always use `https://localhost:3000` for this project's development server, both inside the container and on the container host OS.
- Port 3000 is dedicated to this container. Keep host/container ports at 3000; do not use port 3002 or suggest alternate ports when 3000 is occupied.
- Do not configure `www.localappbuilder.com` or `localappbuilder` hostname aliases, DNS, or hosts-file entries.
- Only local development URLs and local Office manifests use localhost. The deployed Stage origin and endpoints must not change.
- OAuth currently runs on hosted Stage, not localhost. Preserve the pinned Stage callback in `app.config.yaml` and Stage OAuth configuration; do not require localhost OAuth registration or switch OAuth to localhost unless explicitly requested for local OAuth testing.
- Existing `.certs/localappbuilder.crt` and `.certs/localappbuilder.key` filenames are retained for certificate compatibility, not hostname aliasing.
