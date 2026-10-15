# Issue labels

K-SACCR uses the 20-label core catalogue from danielep71/EXCEL-VBA-PROJECT-TEMPLATE at commit b903fe44ef6a032c4689870b83745afa1c22490d.

The label manifest and both reconciliation scripts are copied unchanged. Names, colours and descriptions are defined in `.github/labels.json`. Its `prune: true` policy removes labels outside the selected catalogue during reconciliation, matching the template.

The two workflows are copied from the same revision with two integration adjustments: they pass `--policy .github/label-policy.json` and watch that file instead of the template-wide repository profile, and every checkout sets `persist-credentials: false`, so code under review never receives Git credentials. K-SACCR has not adopted the template's full repository profile. This small policy selects the generated application profile with no domain overlays; all profile overlays are currently empty, so the result is exactly the template's 20 core labels.

- **Sync issue labels** validates pull requests without write permission; pushes to main affecting its inputs and manual runs reconcile live labels and verify an exact match. Only the reconciliation job receives `issues: write`.
- **Detect issue-label drift** checks deterministic fixtures on pull requests and compares live labels daily at 05:17 UTC or on manual request. It is read-only and uploads JSON/Markdown evidence for 30 days. Drift causes failure; detection does not silently repair it.
- Both workflows use the template's full-SHA-pinned actions.

Local checks:

```sh
node .github/scripts/labels-sync.mjs --policy .github/label-policy.json --self-test
node .github/scripts/labels-sync.mjs --policy .github/label-policy.json --mode validate
node .github/scripts/labels-drift.mjs --policy .github/label-policy.json --self-test
```

To update the catalogue, edit the manifest through a pull request. Review removals carefully: deleting a label also removes its issue/PR associations. Labels do not assign priorities or milestones automatically.
