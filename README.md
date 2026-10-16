# EWGCS Resume AI Automation — OpenRouter free-model starter MVP

This project extends the EWGCS Resume Builder with a DOCX upload-and-analyze step. After AI extraction, the result is loaded into the existing editor and live preview. The user reviews and edits the content, then uses the existing Word/PDF export controls. **The application does not export automatically.**

## Live section-by-section preview and review notes

During streaming, the browser shows model token/progress events and a **Live extracted sections** panel. Whenever a source segment has returned valid structured data, the backend sends a cumulative `section_update` SSE event. The frontend appends new section cards and updates previously seen cards using `textContent` so model output is not interpreted as HTML. The live panel can be collapsed after completion to preserve space for editing.

The **Review these AI extraction notes before exporting** panel is displayed immediately below the upload controls and above the editor step navigation. It remains visible on every editor step, has its own scroll area and Show/Hide control, and is excluded from Word/PDF exports. This keeps review notes accessible even when the user is on the final Footer & Export step.

The mapping is source-grounded to prevent section leakage and invented facts: competencies come only from the source Key Expertise/Core Competencies section; certifications come only from Certifications; and dated roles retain their source responsibilities and client details. When a single source paragraph contains many competency phrases, an optional bounded AI call can split it into points. The result is accepted only if it preserves every source word in the same order; otherwise the text is retained and flagged for review. AI summary rewrites are accepted only when their substantive words and numbers occur in the source resume; otherwise a concise source-authored summary is used and a review note is shown. When a resume nests project roles under an employer heading, roles inherit that employer and the employer-level dates/location are retained as labeled role context. Education institution and credential lines are paired without inventing missing graduation years. Skills keep the source's category headings and listed content; when there is no Skills section, the builder leaves skills empty rather than mining unrelated sections. Source section labels such as Objective are retained. Missing employers or partial education details are flagged for review. Duties are never converted into achievements. Empty sections such as Education and Major Achievements are omitted from the final preview/export instead of printing placeholder text.

Long career summaries use a concise source-grounded fallback (up to 110 words) if AI refinement is unavailable; the full source summary remains available in the Summary editor for reference. The experience editor retains every extracted detail, while the preview and export show up to four responsibility points per role by default. Users can switch on all responsibilities before exporting. Multi-line role/company/date layouts and explicit employer/client labels are supported so those details stay in the correct fields.

After import, the editor includes an AI chat panel for plain-language, follow-up edits. It sends the current structured resume and retained source summary/experience context to the configured OpenRouter model, then applies only the returned section updates. If the request fails, the current draft remains unchanged.

## First-version scope

- Input: `.docx` Word resume, up to 10 MB by default.
- Extraction: Python reads document paragraphs, tables, and header/footer text.
- AI: OpenRouter Chat Completions API using an OpenRouter API key. Default model is `openrouter/free`, which routes to an eligible free model supporting JSON output. Model can be changed in `.env`.
- Low-latency mapping: the backend deterministically maps source headings, education, skills, certifications, competencies, clients, and dated role entries. The source summary may be sent for optional concise rewriting using a strict summary-only JSON Schema. A long single-block Core Competencies section may also be sent for boundary-only splitting; the result is accepted only when its words match the source in the same order. Calls use bounded output budgets and timeouts. No entire work-history extraction call is made.
- UI: existing editor and preview, plus **Choose Resume** and **Analyze & Fill** controls, live progress, a source-section preview, and a persistent AI extraction-notes panel.
- Streaming: the backend emits an initial section update from local DOCX parsing, then forwards token progress for optional AI cleanup requests via Server-Sent Events (SSE). The UI can show source-grounded sections immediately, then replace the summary or safely split a long competency block when validation succeeds. If either cleanup request fails, source content is retained and the user gets a review note instead of losing the resume.
- Export: reuses the existing builder's Word-compatible `.doc` export and browser print-to-PDF path.
- The application does not persist resumes or drafts. FastAPI's multipart handling may temporarily spool uploads to the server's temporary directory; configure encrypted ephemeral storage and clean it on the host.

Scanned PDFs, legacy `.doc` files, OCR, authentication, persistent drafts and native `.docx` export are out of scope for this first version.

## 1. Get an OpenRouter API key

1. Sign in or create an account at [OpenRouter](https://openrouter.ai/).
2. Open [OpenRouter API Keys](https://openrouter.ai/settings/keys) and create a key.
3. The key is not a Claude API key. It is used to access models through OpenRouter.
4. Free-model availability and limits can change. `openrouter/free` selects an available model that supports the requested features, but requests may still hit provider or rate limits. See the [OpenRouter free-model router](https://openrouter.ai/openrouter/free).

## 2. Run locally on Windows

Install Python 3.11 or newer, open PowerShell in this folder, and run:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
Copy-Item .env.example .env
notepad .env
```

In `.env`, replace `OPENROUTER_API_KEY=replace_me` with your OpenRouter key. The default model is:

```dotenv
OPENROUTER_MODEL=openrouter/free
OPENROUTER_SUMMARY_MAX_TOKENS=1000
OPENROUTER_SUMMARY_TIMEOUT_SECONDS=45
OPENROUTER_MAX_EXTRACTION_CALLS=48
```

You can set `OPENROUTER_MODEL` to a specific model identifier shown in OpenRouter's model catalog, including a free variant when one is available (often marked `:free`). The summary refinement uses strict JSON Schema response mode and sets `provider.require_parameters=true`, asking OpenRouter to route only to endpoints that support the requested parameter. The default output budget is 1000 tokens and the summary request timeout is 45 seconds; tune `OPENROUTER_SUMMARY_MAX_TOKENS` and `OPENROUTER_SUMMARY_TIMEOUT_SECONDS` if needed. Availability, rate limits and provider behavior still vary. The server logs the actual model selected for the request.

**Keep the key in the backend `.env` file only. Never paste it into `index.html`, commit it to source control, or share it.** `.env` is intended for local development; use a proper secret manager for hosted deployments.

Start the application:

```powershell
python -m uvicorn main:app --reload --host 127.0.0.1 --port 8000
```

Open `http://127.0.0.1:8000`, choose a DOCX resume and click **Analyze & Fill**. Review all imported fields, make any edits, and export Word or PDF.

## API endpoints

- `GET /` — serves the builder.
- `GET /api/health` — local health/configuration check.
- `POST /api/analyze-resume-stream` — accepts multipart field `file` with a `.docx` file and streams status/progress/error events, followed by a `complete` event containing normalized builder data and review notes. The browser uses this endpoint.
- `POST /api/analyze-resume` — non-streaming compatibility endpoint returning normalized builder data plus review notes.

Interactive API docs are available at `http://127.0.0.1:8000/docs`.

## Troubleshooting

- **401**: check the OpenRouter API key and ensure the value has no surrounding quotes/spaces.
- **429**: a free model/provider rate limit may have been reached; wait and retry or select a different available free model.
- **400/404/422 during summary refinement**: confirm the configured model ID and endpoint parameters. Summary rewriting is optional; if the model cannot provide the required structured response, the source summary is retained and a review note is displayed rather than failing the full resume import.
- **`finish_reason="length"` / invalid JSON**: the original source summary or competency text is retained if the model output is invalid or truncated. Summary failures and unsafe competency splits are shown in the review notes. `OPENROUTER_SUMMARY_MAX_TOKENS` controls the summary response budget.
- **Slow/failed requests**: the builder maps job roles and other explicit sections locally, so it does not make full work-history extraction calls. It may make one summary refinement call and one competency-splitting call when the source has one long competency block. Both calls have bounded timeouts; retry or check OpenRouter account limits if a request stalls.

## Accuracy and human review

AI extraction is a drafting aid, not an authoritative source of truth. Confirm names, employers, project/client labels, dates, credentials and every role/responsibility before export. Information missing from the source should be left blank or flagged; the prompt instructs the model not to invent it. A long or table-heavy resume may still require review because Word layouts do not always expose text in a perfectly linear order.

## Before an internal/company rollout

**Current release gate:** this app has no user authentication, SSO, role permissions, or request throttling. Treat it as local development only, or put it behind a company-managed private network and authenticated HTTPS gateway. Do not expose the current API directly to the public internet: anyone who can reach it could upload resumes, incur AI usage, or submit resume-edit requests. A reverse proxy must enforce an HTTP request-body size limit because multipart parsing happens before route-level upload checks.

- Get explicit approval before sending employee resumes to OpenRouter or any selected model provider; resumes can contain personal and confidential employment data. Review the current provider/model routing and data terms. OpenRouter's free router can select different providers over time, so do not assume one uniform retention policy for every free endpoint.
- Add company SSO, role-based access, HTTPS, request throttling, malware scanning, secure secret management and defined file-retention rules. Keep temporary upload storage encrypted and ephemeral, and clean it on the host.
- Keep request logs free of resume contents and API keys.
- Consider native `.docx` generation and a downloadable PDF service if the export format needs to be more deterministic than HTML-as-`.doc` and browser print.
- Add integration tests using sanitized resumes that cover tables, multi-client roles, missing education, multiple jobs, long bullets and malformed files.

## Project files

- `index.html` — existing builder, extended with upload and AI import controls.
- `main.py` — FastAPI backend, DOCX text extraction, OpenRouter call and mapping into the builder's data structure.
- `.env.example` — OpenRouter configuration template; no real secret included.
- `requirements.txt` — Python dependencies.
- `test_resume_automation.py` — basic extraction and data-mapping tests.
