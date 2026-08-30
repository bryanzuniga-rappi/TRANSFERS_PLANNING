#!/usr/bin/env bash
set -euo pipefail

SERVICE_NAME="${SERVICE_NAME:-transfer-planner}"
REGION="${REGION:-us-central1}"
SERVICE_ACCOUNT="${SERVICE_ACCOUNT:-}"

ARGS=(
  run deploy "$SERVICE_NAME"
  --source .
  --region "$REGION"
  --cpu 2
  --memory 4Gi
  --concurrency 1
  --timeout 3600
  --min 0
  --max 3
  --allow-unauthenticated
)

if [[ -n "$SERVICE_ACCOUNT" ]]; then
  ARGS+=(--service-account "$SERVICE_ACCOUNT")
fi

gcloud "${ARGS[@]}"
