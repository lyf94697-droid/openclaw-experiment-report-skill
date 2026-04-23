# Roadmap

## Near Term

- stabilize GitHub-facing repository governance and contribution flow
- improve template compatibility diagnostics
- improve direct-chat reliability when image attachments are staged by the runtime
- strengthen regression fixtures for mixed cover/body templates and grouped image layouts

## Medium Term

- split report-generation logic into reusable document profiles
- move more hardcoded experiment-report assumptions into profile metadata
- support richer profile-specific validation rules
- support reusable style packs beyond `default`, `compact`, and `school`

## Long Term

The repository can support nearby document types only after the experiment-report path stays fast, stable, and easy to verify.

That means:

- one reliable core pipeline for reference gathering, body generation, validation, template filling, image mapping, and final styling
- explicit document-type presets for structure, metadata labels, required sections, caption rules, and validation thresholds
- prompt presets for each document type instead of embedding everything into one giant experiment-report prompt

Potential future profiles:

- internship reports
- course design reports
- training summaries
- project weekly reports
- software test reports
- deployment or operations runbooks
- meeting minutes with template filling

## Non-Goals For Now

- turning the project into a general-purpose office suite
- pretending every document type can be solved by a single catch-all prompt
- shipping a GUI app before the profile model is stable
