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

Clone the repository on the VM first (`git clone https://github.com/lucamattiazzi/sommelier.git`),
then run the following from its root. No container registry is required for this source-build path.

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

### Use an existing reverse proxy

If Caddy already runs in a separate Compose project, use `compose.external-caddy.yaml` on its own
instead of `compose.yaml`. This variant starts only Sommelier, joins Caddy's existing external
network and publishes no host ports. The `expose` setting documents its container port; only Caddy
accepts public traffic. Other containers on the shared network can also reach Sommelier.

In `deploy/sommelier/.env`, set:

```dotenv
SOMMELIER_DOMAIN=sommelier.grokked.it
CADDY_NETWORK=your-existing-network-name
```

Use the actual network name from your Caddy deployment (`docker network ls` lists the networks).
Do not create a second network. From `deploy/sommelier`, validate and start the service:

```sh
docker compose -f compose.external-caddy.yaml config --quiet
docker compose -f compose.external-caddy.yaml up --build -d
docker compose -f compose.external-caddy.yaml ps
```

Add this site block to your **existing Caddyfile**, replacing the domain for another deployment:

```caddyfile
sommelier.grokked.it {
  encode zstd gzip
  reverse_proxy sommelier:3000
}
```

Caddy's [`reverse_proxy`](https://caddyserver.com/docs/caddyfile/directives/reverse_proxy) handles
WebSocket upgrades automatically, so `/connect` needs no separate route. From the directory of
your **Caddy Compose project**, validate and reload (replace `caddy` if its service has another name):

```sh
docker compose exec caddy caddy validate --config /etc/caddy/Caddyfile --adapter caddyfile
docker compose exec caddy caddy reload --config /etc/caddy/Caddyfile --adapter caddyfile
curl --fail https://sommelier.grokked.it/healthz
```

The commands assume your Caddyfile is mounted at `/etc/caddy/Caddyfile`; adapt that path if needed.
Keep access logging disabled for this site, or redact query strings on `/connect`, which carry
relay credentials. Always include `-f compose.external-caddy.yaml` for subsequent Sommelier
updates and lifecycle commands. The external network remains managed by your existing deployment.

### Publish a container image

A registry stores the **built image**, not the Dockerfile. GitHub Container Registry is `ghcr.io`.
Google's former Container Registry (`gcr.io`) has been replaced by Artifact Registry; these are
different services. See the [GitHub registry guide](https://docs.github.com/en/packages/working-with-a-github-packages-registry/working-with-the-container-registry)
and [Google migration guide](https://cloud.google.com/artifact-registry/docs/transition/transition-from-gcr).

To build locally and publish on GHCR, sign in with a GitHub personal access token (classic) with
`write:packages` scope. Enter the token at Docker's password prompt; do not put it in source files.
Run from the repository root, with a running Docker daemon and Buildx:

```sh
docker login ghcr.io -u lucamattiazzi
docker buildx create --name sommelier-builder --driver docker-container --use
docker buildx build \
  --platform linux/amd64,linux/arm64 \
  --file deploy/sommelier/Dockerfile \
  --build-arg SOMMELIER_PUBLIC_ORIGIN=https://sommelier.grokked.it \
  --label org.opencontainers.image.source=https://github.com/lucamattiazzi/sommelier \
  --tag ghcr.io/lucamattiazzi/sommelier:grokked-v1 \
  --push .
```

Create the builder once; reuse it with `docker buildx use sommelier-builder` for later builds.
The two platforms cover both x86 and ARM Hetzner hosts, including builds made on Apple Silicon.
Replace the owner, tag and origin for your own deployment. Use a new version tag for each release.
In GitHub's package settings, make the container package public if it should be downloadable without
authentication: a new GHCR package is private by default, even when the source repository is public.

On the server, replace the entire `build:` block of the `sommelier` service in your chosen Compose
file (`compose.yaml` or `compose.external-caddy.yaml`) with:

```yaml
    image: ghcr.io/lucamattiazzi/sommelier:grokked-v1
```

Keep the service's environment and the Caddy configuration. With `.env` set to
`SOMMELIER_DOMAIN=sommelier.grokked.it`, run:

```sh
docker compose pull
docker compose up -d
curl --fail https://sommelier.grokked.it/healthz
curl --fail -o sommelier-manifest.xml https://sommelier.grokked.it/manifest.xml
```

For an existing external proxy, use `docker compose -f compose.external-caddy.yaml pull` and
`docker compose -f compose.external-caddy.yaml up -d` instead. Point DNS at the VM before checking public URLs.
One image serves the homepage, TaskPane, assets, bridge download and encrypted relay.
The manifest origin is currently embedded **at build time**: changing only the runtime environment
does not rewrite it. Build a new image with the correct origin when self-hosting on another domain.

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

To move to another domain, update `.env` and rebuild with `docker compose up --build -d`. Publish
the new image using the registry workflow if applicable. For an existing proxy, use
`docker compose -f compose.external-caddy.yaml up --build -d` and update your Caddy site block. Distribute
the newly generated manifest and pair again on that instance. Saved associations intentionally keep
their original server address; they are not silently redirected to another operator. Excel stores
saved terminal profiles per TaskPane origin.
