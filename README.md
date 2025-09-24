# ReLead-EDU Deployment Notes

## Apps Script Deployment

This repository is configured to deploy the `apps-script.gs` file to Google Apps Script when changes are pushed to the `main` branch.

### Required configuration

- **Apps Script `scriptId`**: Update [`.clasp.json`](.clasp.json) with the script identifier of the target Apps Script project before running any deployments.
- **`GCP_SA_KEY` secret**: Upload the JSON key for a Google Cloud service account with Editor access to the target Apps Script project as the `GCP_SA_KEY` repository secret.
- **`CLASP_IGNORE` (optional)**: Define this environment variable if you need to override the default `.claspignore` behaviour during automated deployments.
- **Branch protection**: Protect the `main` branch to ensure that linting and review requirements are satisfied before code is merged and deployment is triggered.

### GitHub Actions workflow

The workflow defined at [`.github/workflows/deploy-apps-script.yml`](.github/workflows/deploy-apps-script.yml) performs the following steps:

1. Checks out the repository and installs Node.js dependencies (when a lockfile is present).
2. Optionally runs `npm run lint`, blocking the deployment if lint errors are reported and the script is defined.
3. Authenticates to Google Cloud using `google-github-actions/auth@v1` with the `GCP_SA_KEY` secret and sets up the Cloud SDK via `google-github-actions/setup-gcloud@v1`.
4. Installs the Apps Script CLI (`@google/clasp`) globally and force pushes the local source to Apps Script using `clasp push --force`.

### Manual steps

Some configuration cannot be automated from this repository:

- Creating the Google Cloud service account and downloading its JSON credentials must be done manually in the Google Cloud Console.
- Upload the downloaded JSON as the `GCP_SA_KEY` secret in the GitHub repository settings.
- Configure the required branch protections directly in the repository settings.

## Webhook relay server

The repository now bundles a lightweight Express server that serves the static CRM front-end and exposes a `/webhook` endpoint. Incoming requests are forwarded to the Apps Script deployment (default URL taken from `config/app-config.js`) so Meta can target the Render deployment directly. To run it locally:

```bash
npm install
npm start
```

Set the `APPS_SCRIPT_URL` environment variable if you need to override the Apps Script Web App endpoint.
