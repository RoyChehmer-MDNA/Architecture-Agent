# Skill-driven document agents

Give an agent a task and an image; receive a Word document guided by your own
editable Markdown instructions. The first agent documents software/infrastructure
architecture diagrams. Add other image-to-document agents through `agents.json`.

## Quick start (Windows PowerShell)

Requires Python 3.10 or newer. Install the required packages once:

```powershell
python -m pip install -r requirements.txt
```

1. Edit `skills/architecture.md` with your required structure, terminology, and rules.
2. Choose a provider in the ignored `.env` file. For OpenAI:

```dotenv
AI_PROVIDER=openai
OPENAI_API_KEY=sk-your-key-here
```

Or, to use an existing local AWS login with Amazon Bedrock:

For an SSO-based profile, sign in first with
`aws sso login --profile my-profile`. Access-key profiles do not use this command.

```dotenv
AI_PROVIDER=bedrock
AWS_PROFILE=my-profile
AWS_REGION=us-east-1
BEDROCK_MODEL_ID=us.amazon.nova-pro-v1:0
```

For a default AWS profile, omit `AWS_PROFILE`. `aws login` profiles and other standard
AWS SDK credential sources also work; the requirements include AWS CRT support needed
by Boto3 for `aws login`. Do not put AWS access keys in this project's `.env` file.

3. Copy your architecture image into `input/` (PNG, JPEG, or WebP; up to 20 MiB).
4. Open PowerShell in this project folder and run:

```powershell
python agent_runner.py run --dry-run
python agent_runner.py run
```

The selected account needs access to the configured image-capable model. API calls send
your image, skill text, and task to the selected provider and incur that provider's
usage charges. OpenAI requests use `store: false`; AWS states that Bedrock Converse does
not store submitted text or images. The runner loads `.env` from the project folder.
Values set in the process environment take precedence over values in `.env`.

Open the generated `output/architecture/<run-id>/<source-image-name>.docx` in Word.
For example, `input/system.png` produces a document named `system.docx`.
The document opens with an architecture description followed by the original chart
image, before the detailed sections. The chart fits within the page margins and keeps
its original proportions. PNG, JPEG, and WebP source images are embedded in Word.
Every successful run also saves a copy of the original chart in the same output
folder, preserving its filename and format, plus `skill.md` (the exact instructions
used) and `run.json` (source path, saved chart filename, hashes, task, model, timestamp,
response ID, and usage).
After successful processing, the source image is moved from `input/` to `archive/`.
An image that fails remains in `input/` so it can be retried. Dry runs never move files.
If an archived filename already exists, the runner adds a timestamp and unique suffix
instead of overwriting either image. Existing documents are preserved.

## Automatically process the input folder

```powershell
python agent_runner.py
```

Starting the runner without a command defaults to watch mode. The explicit command
`python agent_runner.py watch` works too. While running, it automatically picks up
new images added to `input/`, generates documents using the current skill file, and
moves completed images to `archive/`.

Leave this terminal running and press Enter when you want it to stop. Ctrl+C remains
available as a fallback. If Enter is pressed during an API request, the current document
finishes before the process exits. The watcher waits for an image's size and modification
time to remain unchanged across two polls. For large/slow copies,
copy using a temporary extension and rename to `.png`/`.jpg`/`.webp` when finished.
Run only one watcher per agent to avoid duplicate API calls.

Unchanged successful jobs are skipped, including after restarting. Editing the skill,
task, model, or image creates a new job on the next scan. Editing the skill while a
request is running applies to the next request. Failed API jobs are attempted once per
watcher session to avoid repeated charges; fix the issue and restart to retry.

## Give a custom task

```powershell
python agent_runner.py run --image input/system.png --task "Document the design for an operations handover"
python agent_runner.py run --force
python agent_runner.py list
```

`--force` regenerates unchanged images. `--dry-run` validates file/configuration loading
without calling the API; it does not check image decoding, account access, or output quality.
Folder processing creates one independent document per image; images are not combined.

## Choose the document language and direction

Set `DOCUMENT_LANGUAGE` in `.env`. This runtime value overrides the fallback language
written in the skill file:

```dotenv
DOCUMENT_LANGUAGE=IL
```

`IL` and `HE` select Hebrew. `AR`, `FA`, `UR`, and `YI` are also treated as RTL
languages. All Word headings are centered. For RTL languages, body paragraphs, lists,
code blocks, and table cells are right-aligned, with right-to-left reading direction
also applied to headings and table direction, including lines that contain English
technical identifiers. Common LTR values include `EN`, `FR`, `ES`, and
`DE`. Restart the watcher after changing `.env`.

## Choose heading appearance

Set `heading_style` inside each agent's entry in `agents.json`:

```json
"heading_style": {
  "color": "#000000",
  "font": "Arial",
  "size": 16,
  "bold": true,
  "italic": false,
  "underline": false
}
```

These settings apply to all Word headings, including the document heading. `color`
accepts a six-digit hex value with or without `#`; `font` is a font name installed
on the computer opening the document; `size` is in points, from 1 to 96. Omitted
settings default to black, Arial, and 16 points. `bold`, `italic`, and `underline`
accept JSON `true` or `false`; their defaults are `true`, `false`, and `false`,
respectively. These settings override inline Markdown formatting in headings and
apply to RTL text too. Headings remain centered and keep
their language's reading direction. The watcher reads changes on its next scan.
Previously generated documents keep their existing appearance.

## Choose a provider and model

`AI_PROVIDER` accepts `openai` (the default) or `bedrock`. OpenAI needs
`OPENAI_API_KEY`; Bedrock does not read or require it. Bedrock uses Boto3's standard
credential chain, including AWS CLI/SDK profiles and cached IAM Identity Center (SSO)
sessions. Your identity needs `bedrock:InvokeModel` permission for the selected model.

For Bedrock, set an image-capable model or inference profile that supports the Converse
API:

```dotenv
AI_PROVIDER=bedrock
AWS_PROFILE=my-profile
AWS_REGION=us-east-1
BEDROCK_MODEL_ID=us.amazon.nova-pro-v1:0
# Optional response limit:
BEDROCK_MAX_TOKENS=8192
```

The model must be available in the chosen region and enabled for your AWS account.
Amazon Nova Pro is shown as an example; use the model/inference-profile ID appropriate
to your region and account. If your SSO session expires, run
`aws sso login --profile my-profile` again. The runner never launches an interactive
login itself.

For `eu-west-1`, these image-capable Converse profiles were verified as active for this
project's AWS account. Costs below estimate one document with the latest observed usage
of 3,274 input tokens (including the image and instructions) and 1,760 output tokens.
They are estimates based on published standard on-demand token prices and exclude taxes,
currency conversion, and any future AWS pricing changes.

| Cost order | `BEDROCK_MODEL_ID` | Best use | Estimated document cost |
| --- | --- | --- | ---: |
| Cheapest | `eu.amazon.nova-lite-v1:0` | Simple diagrams and high volume; review carefully for omissions | about `$0.00062` |
| Recommended | `eu.amazon.nova-pro-v1:0` | Best cost/quality balance for architecture and security reviews | about `$0.0083` |
| Strong/fast | `eu.anthropic.claude-haiku-4-5-20251001-v1:0` | Strong instruction following with modest cost | about `$0.0133` |
| Higher quality | `eu.anthropic.claude-sonnet-4-6` | Complex diagrams and deeper gap analysis | about `$0.040` using comparable EU Sonnet geo pricing |
| Premium | `eu.anthropic.claude-opus-4-5-20251101-v1:0` | Difficult diagrams where maximum review depth matters | about `$0.066` using comparable EU geo pricing |

Change only the active value and restart the watcher:

```dotenv
# Cheapest
BEDROCK_MODEL_ID=eu.amazon.nova-lite-v1:0

# Recommended default
# BEDROCK_MODEL_ID=eu.amazon.nova-pro-v1:0
```

The actual cost formula is
`(input tokens / 1,000,000 × input price) + (output tokens / 1,000,000 × output price)`.
Each run stores its real token counts under `usage` in `run.json`. Image complexity,
resolution, prompt length, language, and response length can change the token count.

For OpenAI, set `OPENAI_MODEL` in `.env` to an image-capable Responses API model:

| Value | When to use it |
| --- | --- |
| `gpt-6-astra` | Highest quality for difficult diagrams and detailed architecture reasoning |
| `gpt-5.6-sol` | Strong general-purpose professional work |
| `gpt-5.6-terra` | Balance of quality and cost |
| `gpt-5.6-luna` | Lowest-cost option for simpler diagrams or high volume |

For example:

```dotenv
OPENAI_MODEL=gpt-5.6-terra
```

Restart the watcher after changing `.env`. You can instead change `model` in
`agents.json` or override it for one run:

```powershell
python agent_runner.py run --model gpt-5.6-luna
```

Model selection priority is `--model`, then `BEDROCK_MODEL_ID` or `OPENAI_MODEL` for
the selected provider, then the value in `agents.json`. Since the registry's starter
model is an OpenAI model, set `BEDROCK_MODEL_ID` whenever Bedrock is selected. The
selected model must be available to your provider account. Consult the
[current OpenAI model catalog](https://developers.openai.com/api/docs/models) for updated
availability and the
[OpenAI image input documentation](https://developers.openai.com/api/docs/guides/images-vision)
for OpenAI details. For Bedrock, see the
[model availability and compatibility guide](https://docs.aws.amazon.com/bedrock/latest/userguide/models.html)
and the [Converse API reference](https://docs.aws.amazon.com/bedrock/latest/APIReference/API_runtime_Converse.html).

## Maintain your skill

`skills/architecture.md` is ordinary Markdown, with no special syntax required.
Replace or extend its starter instructions whenever your requirements change.
Only this file's contents are loaded; referenced files are not automatically included.
No retraining or restart is needed for skill updates.

The agent is prompted to check compliance and distinguish observed facts from assumptions.
Architecture reports also include explanatory bullets that walk through visible data
flows, discuss design tradeoffs and hypothetical failure or growth scenarios, and give
readers focused questions and practical validation steps.
This is model-guided compliance, not a deterministic guarantee. Review the generated document,
especially small labels, arrow directions, and inferred details.

## Add another agent

Create `skills/security-review.md`, create `input/security/`, and add an entry to `agents.json`:

```json
"security-review": {
  "skill": "skills/security-review.md",
  "input": "input/security",
  "archive": "archive/security-review",
  "output": "output/security-review",
  "task": "Review the diagram and produce a security assessment.",
  "model": "gpt-6-astra"
}
```

Run `python agent_runner.py run --agent security-review`.
Paths in the registry are relative to the project folder; `--image` is relative to your
current terminal directory. This starter supports configurable image-to-Word agents;
other input types, generated applications, and chained workflows require additional code.

## Verification

```powershell
python -B -m unittest discover -s tests -v
```

Tests use simulated provider calls and cover successful output, skill snapshots, change
detection, duplicate skipping, dry runs, and failure handling. A real image and provider
login are required to verify live generation.

# OpenAI API Key link
# Revoke it at OpenAI API keys.
