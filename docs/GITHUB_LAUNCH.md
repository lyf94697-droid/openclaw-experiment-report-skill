# GitHub Launch Notes

Use this file when preparing the repository page, first release, and social preview on GitHub.

## About Section

Recommended repository description:

```text
OpenClaw skill and PowerShell pipeline for generating Chinese university lab reports, filling docx templates, embedding screenshots, and producing polished final documents.
```

Recommended topics:

```text
openclaw
skill
powershell
docx
report-generation
template-filling
document-automation
lab-report
openxml
university
```

## Social Preview

Recommended options:

1. Use a screenshot of the top section of `README.md` after badges and preview images load.
2. Use a cropped screenshot from `demo/README.md` that shows the 2x2 layout preview.
3. Use a manually composed image built from:
   - `demo/assets/step-network-config.png`
   - `demo/assets/step-ipconfig.png`
   - `demo/assets/result-ping.png`
   - `demo/assets/result-arp.png`

Suggested overlay text for a preview image:

```text
OpenClaw Experiment Report Skill
Generate Chinese lab reports, fill docx templates, and insert grouped screenshots.
```

## First Release

Recommended tag:

```text
v0.1.0
```

Recommended release title:

```text
v0.1.0 - Initial open-source release
```

Recommended release notes:

```markdown
## Highlights

- Generate Chinese university experiment report bodies from topics, screenshots, tutorial URLs, and structured requirements.
- Fill blank Word or docx templates through deterministic field mapping.
- Insert grouped screenshots with captions, including 2x2 layout blocks anchored under one section.
- Apply final formatting for headings, metadata rows, body text, captions, and image groups.
- Support chat-friendly local wrappers for Feishu-style direct use.

## Included

- OpenClaw skill instructions
- PowerShell helper scripts
- Demo assets and prompt examples
- Smoke tests and GitHub workflows
- Contribution, security, support, and roadmap files

## Known Scope

- Designed for OpenClaw-based local workflows
- Best tested on Windows and PowerShell
- Focused on experiment-report pipelines, with future expansion planned through document profiles
```

## Post-Publish Checklist

- Confirm the `About` description and topics are set.
- Upload a social preview image.
- Verify both GitHub Actions workflows appear and run on the default branch.
- Create the first `v0.1.0` release from tag.
- Pin the repository on your GitHub profile if this is one of your main public projects.
