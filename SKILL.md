---
name: prepare-cn-online-court-materials
description: Extract facts from user-provided chats, transaction records, identity clues, screenshots, audio, video, documents, and platform records; triage suspected online small-value scams between civil filing, police reporting, or human review; organize electronic evidence; and generate editable DOCX plus upload-ready PDF materials for People's Court Online Service in China. Use to prepare a civil complaint for small-value online goods transactions between individuals, or to create routing and evidence packets for online services, deposits, private transfers, loans, investment/rebate schemes, non-delivery, false-description, refund, and similar disputes. Do not use for representation, guaranteed legal conclusions, automatic filing, or high-complexity criminal, investment, virtual-currency, cross-border, class/mass-victim, or urgent-preservation matters without professional review.
---

# Prepare Chinese Online Court Materials

Create a simple, evidence-first workflow. Extract before asking, ask only blocking questions, and produce court-ready drafts without inventing facts or legal conclusions.

## Non-negotiable interaction rules

1. Inspect every supplied file and prior answer before asking anything.
2. Reuse facts that can be extracted reliably. Never ask the user to retype visible information.
3. Ask only when a missing fact changes routing, party identity, jurisdiction, claims, evidence meaning, signature, or final-submission status.
4. Ask exactly one question at a time.
5. Prefer a closed choice with 2-4 plain-language options. Include `暂不确定` when uncertainty is legitimate. Use a free-form question only when no safe finite choice exists.
6. Recommend the simplest safe default and explain its consequence in one short sentence.
7. Separate `材料中明确显示`, `用户确认`, `合理推测`, and `仍缺失`. Never promote a guess into a pleaded fact.

## Workflow

### 1. Preserve and inventory

- Work on copies. Tell the user to retain original devices, source files, platform records, and downloadable bills.
- Create a file inventory with original filename, type, size, visible account identity, date range, and extraction status.
- Detect duplicates and near-duplicates. Keep the most complete version; do not silently discard originals.
- Do not upload personal data to an unrelated external service. Redact working previews where possible, but preserve an unredacted court-submission copy under user control.
- Read [evidence-rules.md](references/evidence-rules.md) before processing screenshots, chats, payment records, webpages, audio, or video.

### 2. Extract into structured case data

- Populate the schema in [case-data.example.json](assets/case-data.example.json).
- Extract parties, account handles, phone numbers, addresses, transaction time, amount, payment channel, order information, promises, performance, reminders, refund discussions, platform complaints, police contact, and losses.
- Store a source pointer for every extracted fact: filename plus page, image, timestamp, or message range.
- Mark confidence as `confirmed`, `user_confirmed`, `inferred`, or `missing`.
- Read [intake-and-routing.md](references/intake-and-routing.md) for the minimum data and question order.

### 3. Route before drafting

Version-one scope rule: generate a civil complaint only when `case_type` is `online_goods` and the transaction is a goods sale between individuals. For services, loans, investment/rebate schemes, virtual currency, account rental, or any other transaction type, generate only routing, preservation, and evidence-organization materials; use `police` or `review`, not `civil`. Do not bypass this rule merely because the user requests a complaint.

Return one provisional route:

- `civil`: a genuine transaction or civil relationship is identifiable, a defendant can be distinguished from others, the user has a concrete civil request, and a plausible court/jurisdiction basis can be checked.
- `police`: the material indicates possible deception rather than an ordinary performance dispute, identity or account use may be false or stolen, there may be multiple victims, or the scheme involves investment, virtual currency, account rental, task-rebate, impersonation, or similar criminal-risk signals.
- `review`: civil/criminal characterization, defendant identity, jurisdiction, limitation, platform liability, evidence authenticity, or claim design is materially uncertain.

Do not declare that conduct constitutes fraud or that a case will be accepted. Describe the route as a procedural recommendation based on current materials. A police report does not automatically bar preservation of a possible civil claim.

### 4. Verify current official rules

- Before producing a final-submission version, verify material legal and platform rules against official government, legislature, procuratorate, public-security, or court websites.
- Never rely on a commercial legal article when an official primary source exists.
- Record the issuing body, document title, URL, effective/version date when available, and verification date.
- Start with [official-sources.md](references/official-sources.md). Treat it as a routing index, not a frozen statement of current law.
- If the People's Court Online Service screen shows local file-size, format, naming, or category requirements, follow the live platform notice over a generic rule.

### 5. Build the requested packet

Read [document-spec.md](references/document-spec.md), then generate:

For `civil`:

- filing-readiness report;
- civil complaint using the applicable 2025 Supreme People's Court element-style template where available;
- party information sheet when needed;
- evidence index;
- numbered evidence packet;
- amount calculation sheet when needed;
- upload order and final confirmation checklist.

For `police`:

- route explanation and urgent preservation list;
- factual report statement without declaring guilt;
- loss and payment table;
- suspect account and platform clue table;
- evidence index and numbered evidence packet;
- records that investigators may request from platforms or payment institutions.

For `review`:

- do not generate a final complaint;
- generate a readiness report, missing-items list, preserved-evidence index, and the single next closed-choice question.

Use `scripts/build_case_packet.py` after case data is normalized:

```bash
python scripts/build_case_packet.py case.json --output-dir packet --mode draft
python scripts/build_case_packet.py case.json --output-dir packet --mode final
```

Generate both DOCX and PDF for authored documents. Generate the combined evidence materials as PDF. Keep each upload unit separate and use filenames beginning with `01-`, `02-`, and so on.

### 6. Enforce the final gate

Refuse `--mode final` when any of the following remains:

- route is `review`;
- a pleaded fact or amount is only inferred;
- plaintiff identity or signature date is missing;
- civil route lacks a distinguishable defendant, concrete claim, essential facts, or court confirmation;
- civil route is not an `online_goods` transaction between individuals;
- evidence index lacks source or proof purpose;
- a critical open issue remains;
- live platform upload requirements have not been checked.

Place `草稿-待核对` prominently on draft files. Never sign for the user, simulate a seal, submit to a court, send to police, pay fees, or make an irreversible filing action.

### 7. Render and inspect

- Render every generated DOCX to PNG pages and inspect every page.
- Render every generated PDF to PNG pages and inspect every page.
- On macOS, if LibreOffice renders Chinese as empty boxes, rerender with `FONTCONFIG_FILE=assets/fontconfig-macos.conf`. This exposes the system CJK fonts to the isolated renderer; do not treat missing preview glyphs as missing document content.
- Check Chinese glyphs, tables, page numbers, image orientation, cropping, duplicates, blank pages, evidence numbering, and index-to-page consistency.
- Rebuild after any defect. Do not deliver files that have only passed text extraction.

## Safety boundaries

- Present outputs as self-help drafts, not individualized legal representation.
- Do not promise filing acceptance, case characterization, victory, recovery, police registration, or enforcement.
- Do not fabricate identity, address, jurisdiction, dates, interest, damages, statutes, evidence, quotations, signatures, or seals.
- Do not advise editing original evidence. Put annotations only on copies and label them.
- Escalate to a lawyer, court litigation service center, public-security organ, or other competent body for urgent limitation, preservation, unknown identity, cross-border facts, minors, incapacity, investment/virtual-currency schemes, many victims, platform co-liability, or significant loss.

## Reference map

- Read [intake-and-routing.md](references/intake-and-routing.md) for extraction, question order, and route signals.
- Read [evidence-rules.md](references/evidence-rules.md) for electronic-evidence handling.
- Read [document-spec.md](references/document-spec.md) for file contents, naming, and evidence-index fields.
- Read [official-sources.md](references/official-sources.md) whenever legal or platform rules affect the output.
