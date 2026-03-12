## Commit Message Instructions

When generating commit messages, always append the following trailers on new
lines after the commit body:

```text
Co-authored-by: GitHub Copilot <copilot@github.com>
Co-authored-by-model: <model-name>
```

Replace `<model-name>` with the exact AI model that assisted. Examples:

- `Claude Opus 4.6`
- `Claude Haiku 4.5`
- `GPT-4o`
- `GPT-4.1`
- `Gemini 2.5 Pro`

### Format

```text
<type>: <short summary>

<optional body>

Co-authored-by: GitHub Copilot <copilot@github.com>
Co-authored-by-model: Claude Opus 4.6
```

### Rules

- Always include the `Co-authored-by` trailer — it must be the second-to-last
  line.
- Always include the `Co-authored-by-model` trailer — it must be the last line.
- Separate the trailers from the body with a blank line.
- Do not duplicate either trailer if it is already present.
- Use the exact model name as reported by the AI tool (do not abbreviate).
