# Self-host Sommelier

The deployment runs the Sommelier task pane and authenticated WebSocket relay behind Caddy. Caddy obtains
and renews the public TLS certificate. The Node process is reachable only inside the Compose network.
Deploy once on Hetzner or another Docker host. These instructions are for the service operator;
end users only install the add-in and local harness adapter.

`SOMMELIER_DOMAIN` is the deployment setting, not a domain compiled into the product. Set it to your own
DNS name. Compose uses it for Caddy, the server's public origin and the generated Office manifest.
The TaskPane discovers its relay through its hosting origin; pairing carries that address to the
adapter. No shared registry, account or dependency on another Sommelier operator is required.

## VM prerequisites

- A Linux VM with Docker Engine and the Compose plugin.
- An `A`/`AAAA` DNS record pointing the Sommelier domain to the VM.
- Inbound firewall access only to TCP 22, 80, 443 and UDP 443.

## Deploy

```sh
cd deploy/sommelier
cp .env.example .env
# Set SOMMELIER_DOMAIN in .env, then:
docker compose up --build -d
docker compose ps
docker compose exec sommelier node -e "fetch(process.env.SOMMELIER_PUBLIC_ORIGIN + '/healthz').then(r => { if (!r.ok) process.exitCode = 1; return r.text(); }).then(console.log)"
```

Do not put agent API keys in `.env`. Direct BYOA credentials remain in the task pane's memory and go
straight to the configured agent endpoint. E2EE keys and saved associations remain on the endpoints;
the relay keeps only an in-memory routing registry and forwards encrypted frames. Saved profiles
survive relay restarts and reconnect without a new association. Caddy access logging is intentionally not enabled because relay credentials
are carried in WebSocket query strings.

## Office manifest

The Docker build already generates and serves `/manifest.xml` for `SOMMELIER_DOMAIN`. Download it from
your instance and sideload it in Excel. For a separate source build, generate with the same origin:

```sh
SOMMELIER_PUBLIC_ORIGIN=https://sommelier.example.com \
  pnpm --filter @lucamattiazzi/sommelier-addin manifest:production
```

The production file is `apps/addin/public/manifest.xml`. The build serves it at `/manifest.xml`;
the separate `apps/addin/manifest.xml` remains the localhost development manifest.
Share the production XML from your instance or repository for manual installation in Excel.

The same Node server serves the static English home page at `/`, the TaskPane at
`/taskpane.html` and encrypted WebSockets at `/connect`. No additional web server application
or frontend deployment is needed; Caddy handles public HTTPS.

For the project's hosted instance, set `SOMMELIER_DOMAIN=sommelier.grokked.it` in `.env`.
Create the domain's DNS record before starting Caddy. The committed hosted manifest is generated
with `pnpm manifest:hosted`; self-hosted builds always use their own `SOMMELIER_DOMAIN`.

## Change instance

To move to another domain, update `.env` and rebuild with `docker compose up --build -d`. Distribute
the newly generated manifest and pair again on that instance. Saved associations intentionally keep
their original server address; they are not silently redirected to another operator. Excel stores
saved terminal profiles per TaskPane origin.
