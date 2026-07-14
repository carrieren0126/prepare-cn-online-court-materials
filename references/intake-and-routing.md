# Intake and routing

## Extraction order

Extract before asking:

1. User identity and contact details from supplied identity materials.
2. Counterparty name/entity, address clues, account handles, phone, payment account, platform ID, and delivery details.
3. Transaction subject, amount, payment channel, time, promises, delivery/performance, reminders, refund discussion, and loss.
4. Platform complaint, payment dispute, police contact, mediation, arbitration, or prior litigation.
5. Evidence provenance: file, page/image/message range, original carrier, and whether a platform-generated record is available.

Do not infer an identity merely because the same display name appears across platforms.

## First unanswered question order

Ask only the highest item still blocking progress:

1. Immediate preservation or limitation risk.
2. Transaction type: goods purchase / paid service / loan or private lending / investment or rebate task / other / unsure.
3. Desired route when materials support more than one: civil recovery / police report / preserve both / unsure.
4. Counterparty identity sufficiency.
5. Concrete remedy: refund/return / payment / performance / compensation / unsure.
6. Court or jurisdiction confirmation.
7. Mediation preference.
8. Signature and submission date.

Use a closed choice whenever possible and ask one question only.

## Version-one product scope

- `online_goods`: may enter the `civil` route and generate a civil complaint when all other gates pass.
- `online_service`, `private_lending`, `investment_or_rebate`, `virtual_currency`, `account_rental`, and `other`: do not generate a civil complaint. Use `police` or `review` and produce only routing, preservation, missing-items, and evidence-organization materials.
- If the transaction type is unclear, ask one closed-choice question before selecting `civil`.
- The individual-to-individual condition must be visible in the materials or confirmed by the user; otherwise use `review`.

## Provisional route signals

### Civil-leaning within version one

- A real goods-sale relationship between individuals existed.
- The dispute centers on non-delivery, quality, refund, payment, or performance.
- The counterparty can be distinguished by name/entity plus address or equivalent specific information.
- The user can state a concrete civil remedy and essential facts.

### Police-leaning

- No genuine performance appears intended from the outset.
- Identity, storefront, account, or credentials may be false, stolen, rented, or rapidly replaced.
- The pattern includes impersonation, investment, virtual currency, task rebate, account rental, advance-fee unlocking, or repeated additional payments.
- Multiple victims or many receiving accounts appear.
- Platform or payment data likely requires investigative authority to obtain.

### Human-review triggers

- Defendant cannot yet be distinguished.
- Civil and criminal facts overlap materially.
- Limitation, jurisdiction, arbitration, platform liability, preservation, or service address is uncertain.
- The user seeks punitive damages, emotional damages, interest, attorney fees, or other amounts without a clear basis.
- The matter is cross-border, involves a minor/incapacitated person, virtual assets, many victims, significant loss, or urgent asset/evidence preservation.

## Route language

Say: `根据当前材料，建议优先准备……；最终是否受理或立案由有权机关依法审查。`

Do not say: `这就是诈骗`, `法院一定会立案`, `一定能追回`, or `报警没有用`.
