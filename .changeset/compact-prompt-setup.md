---
"@lucamattiazzi/sommelier-addin": patch
"@lucamattiazzi/sommelier-server": patch
---

Make the initial TaskPane show a ready-to-copy agent prompt, with no required terminal name or
preinstalled skill. The prompt bootstraps the published local bridge and includes workbook RPC
instructions. Keep the connected chat and composer in view, move secondary controls into Options,
and allow renaming the default connection after pairing.

Allow development relay and TaskPane proxy ports to be configured together with
SOMMELIER_RELAY_PORT. Update setup documentation and cover prompt copying, clipboard fallback,
renaming, approvals and remembered reconnection in the narrow-pane browser test.
