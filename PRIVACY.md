# Airy Privacy

Last updated: September 10, 2026

Airy opens, edits, and saves documents locally. Document editing does not
upload files anywhere. AI features require a network connection and send
requests only when you use them.

## No usage analytics

This fork ships without usage analytics: no analytics code runs, no telemetry
endpoints are contacted, and no usage events of any kind are collected or sent.

## AI features

AI features run only against providers you explicitly configure (an API key of
your own or a local endpoint). Keys are stored locally and sent only to the
provider you chose. Document content is sent to that provider only when you
invoke an AI feature on it.
