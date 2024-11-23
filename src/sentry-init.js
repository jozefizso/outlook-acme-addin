import { useEffect } from "react";
import * as Sentry from "@sentry/react";

Sentry.init({
  dsn: 'https://b4ff88695769edca23b38d000560dd77@o4506418945392640.ingest.us.sentry.io/4508347524644864',
  release: 'outlook-acme-addin@0.1.0',
  environment: import.meta.env.VITE_ENVIRONMENT,
  integrations: [
    Sentry.browserTracingIntegration(),
  ],
  tracesSampleRate: 1.0,
  tracePropagationTargets: ['localhost', /^https:\/\/outlook-acme-addin\.pages\.dev/],
})
