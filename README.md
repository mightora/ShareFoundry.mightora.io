# ShareFoundry by Mightora

Browser-first tools for SharePoint makers. The first release provides a list schema designer, Microsoft Graph create-list request JSON, schema CSV, local project import/export, and a column formatting studio with previews.

Live URL (after GitHub Pages and DNS are configured): https://sharefoundry.mightora.io

## Develop

Requires Node.js 20+. Run `npm run check`, then serve `dist/` using any static HTTP server or deploy its contents to GitHub Pages. There are no runtime dependencies for the application core.

## Deployment

- GitHub Pages: use the provided `pages.yml` workflow after publishing this source in a GitHub repository.
- The root `CNAME` file is copied into `dist/` at build time for the GitHub Pages custom domain `sharefoundry.mightora.io`.
- Offline: open `dist/offline/sharepoint-maker-toolkit.html` directly from disk, or download the ZIP from the online site.
- A GitHub Actions test workflow and CodeQL workflow are included. Do not display passing badges until they have run successfully on the target public repository.

## Security and limitations

- No tenant sign-in or write actions. JSON export is not deployment.
- No user data is sent to a Mightora application backend. The online edition loads shared UI libraries from CDNs; the offline edition uses no network resources.
- LocalStorage keeps a draft when browser policy permits. Export project JSON to transfer/backup it.
- Microsoft Graph supports a much wider collection of column features than this initial designer. Current export supports text, number, Yes/No, date-only, and choice; test the request in the target tenant before adoption.
- Generated formatters must be tested against actual list column types and values. Progress examples assume numeric 0–100; URL formatting assumes hyperlink/text values containing safe intended URLs.
- This is a design and generation aid, not a substitute for SharePoint access review, production change control or tenant validation.

## Roadmap

- Editable existing columns, list views and index patterns.
- SharePoint site templates, relationship mapping, library designer.
- Optional delegated Microsoft sign-in for explicit, audited deployments.
- Schema compare, migration assistance and permission explorer.

Copyright (c) 2026 Mightora / Ian Tweedie. MIT licensed.
