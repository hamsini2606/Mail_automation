# 📬 Bulk Cold Email Automation — Internship Outreach

A production-ready Python script for sending personalized, spam-resistant cold emails at scale. Built for internship outreach — sends a polished HTML email with your resume attached to every contact in your Excel sheet, with full logging, retry logic, and multiple anti-spam measures baked in.

---

## 📁 Project Structure

```
.
├── send_emails.py          # Main script — run this
├── Opening_Lines.xlsx      # 96 unique opening lines + subject variants (cycles on loop)
├── sample.xlsx             # Sample contacts file — see this to understand the format
├── Mihir_Kulkarni_Resume.pdf
├── .env                    # Your Gmail credentials (never commit this)
├── .gitignore
├── requirements.txt
│
├── sent_emails.xlsx        # Auto-generated after run (gitignored)
├── failed_emails.xlsx      # Auto-generated after run (gitignored)
└── skipped_emails.xlsx     # Auto-generated after run (gitignored)
```

---

## ⚙️ Setup

### 1. Clone the repo

```bash
git clone https://github.com/Mihirkool/your-repo-name.git
cd your-repo-name
```

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

**requirements.txt:**
```
pandas
openpyxl
python-dotenv
email-validator
```

### 3. Create your `.env` file

```bash
touch .env
```

Add your Gmail credentials inside it:

```env
EMAIL_ADDRESS=your.email@gmail.com
EMAIL_PASSWORD=your_app_password_here
```

> ⚠️ **Do NOT use your regular Gmail password.** You need a Gmail App Password.
> Go to: [Google Account → Security → 2-Step Verification → App Passwords](https://myaccount.google.com/apppasswords)
> Generate one for "Mail" and paste it here.

### 4. Prepare your contacts file

Rename your contacts Excel file to match `CONTACTS_FILE` in the script (default: `contacts 3.xlsx`).

Your Excel file **must have these exact column names** (case-insensitive):

| name | designation | company | mail |
|------|-------------|---------|------|
| Riya Sharma | HR Manager | Zepto | riya@zepto.com |
| Arjun Mehta | CTO | Razorpay | arjun@razorpay.com |

See `sample.xlsx` for the exact format.

### 5. Run

```bash
python send_emails.py
```

---

## 🎛️ Configuration

All tunable settings are at the top of `send_emails.py`:

```python
CONTACTS_FILE      = "contacts 3.xlsx"      # Your contacts Excel file
RESUME_FILE        = "Mihir_Kulkarni_Resume.pdf"
VARIANTS_FILE      = "Opening_Lines.xlsx"   # Opening lines + subject pool

MIN_DELAY          = 35       # Min seconds between emails
MAX_DELAY          = 75       # Max seconds between emails (random jitter)
BATCH_SIZE         = 20       # Emails per batch before a long break
BATCH_BREAK_MIN    = 8 * 60   # Min cool-down between batches (seconds)
BATCH_BREAK_MAX    = 15 * 60  # Max cool-down between batches (seconds)
MAX_EMAILS_TO_SEND = 500      # Hard cap — set to 10 for testing
MAX_RETRIES        = 3        # Retry attempts per failed email
```

**For testing**, set `MAX_EMAILS_TO_SEND = 5` and send to your own email first.

---

## 📊 Opening_Lines.xlsx

This file contains **96 unique variations** of the email opening paragraph and subject line. Every email sent picks the next row in sequence — so no two consecutive emails ever look identical to spam filters.

| Column | Description |
|--------|-------------|
| **Opening Lines** | First paragraph of the email body |
| **Subject** | Email subject line for that variant |

### Placeholder system

Both columns support dynamic placeholders that get resolved at send time:

```
{name}          →  Contact's name        e.g. "Riya"
{designation}   →  Their job title       e.g. "HR Manager"  
{company}       →  Their company name    e.g. "Zepto"
```

**Example row from the file:**

| Opening Lines | Subject |
|---|---|
| Your profile as `{designation}` at `{company}` stood out to me... | Internship Inquiry – Mihir Kulkarni \| AI/ML & Full-Stack |

**At send time this becomes:**

> *Your profile as HR Manager at Zepto stood out to me...*

### Infinite cycling

The script uses `itertools.cycle()` to loop through all 96 rows endlessly:

```
Email  #1  → Row 1
Email  #2  → Row 2
...
Email #96  → Row 96
Email #97  → Row 1  (loops back)
Email #98  → Row 2
...and so on forever
```

This means even if you have 500+ contacts, every email still gets a unique opening relative to the emails around it in the batch.

To add more variants, just add rows to `Opening_Lines.xlsx` — no code changes needed.

---

## 🛡️ Anti-Spam Measures

Getting cold emails into the inbox (not spam) requires looking like a human, not a bot. Here's every measure implemented and *why* it works:

### 1. Content Fingerprint Variation

Spam filters compute a hash of your email's content. If 500 identical emails go out from one sender, every major filter (Gmail, Outlook, SpamAssassin) flags it as bulk.

**What we do:** Every email has a different opening paragraph and subject line pulled from `Opening_Lines.xlsx`. Combined with the personalized `{name}`, `{designation}`, and `{company}` values, no two emails share the same content fingerprint.

---

### 2. Plain Text Alternative (multipart/alternative)

HTML-only emails are a major spam signal. Real personal emails almost always have a plain text version alongside the HTML — it's how email clients were designed to work.

**What we do:** Every email is sent as `multipart/alternative` with both a plain text body and the HTML body. Spam filters see this and treat it as a legitimate dual-format email.

```python
alt = MIMEMultipart("alternative")
alt.attach(MIMEText(plain_body, "plain"))   # plain text first
alt.attach(MIMEText(html_body,  "html"))    # HTML second (clients prefer last match)
```

---

### 3. Random Delay with Jitter (Human Typing Pattern)

Sending emails at a fixed interval (e.g. exactly every 8 seconds) is an immediate bot signature — humans don't behave that precisely.

**What we do:** A random delay between `MIN_DELAY` and `MAX_DELAY` seconds (default 35–75s) is inserted between each email. This creates an irregular, human-like sending pattern.

```python
delay = random.randint(MIN_DELAY, MAX_DELAY)
time.sleep(delay)
```

---

### 4. Batch Cooling Breaks

Gmail's spam system monitors sending velocity. Sending 500 emails in a single session, even slowly, can trigger a bulk-send flag and eventually suspend your account.

**What we do:** Every `BATCH_SIZE` emails (default: 20), the script pauses for a random `BATCH_BREAK_MIN` to `BATCH_BREAK_MAX` duration (default: 8–15 minutes), then reconnects SMTP and continues. This mimics a human taking breaks between sessions.

```
Send 20 emails → pause 8-15 mins → reconnect → send 20 more → ...
```

---

### 5. Subject Line Variation

Identical subject lines to hundreds of recipients is one of the oldest and most reliable spam signals. Filters flag it even if the body content differs.

**What we do:** Subject lines are also pulled from `Opening_Lines.xlsx` and rotate with the opening lines. Some subjects also include `{company}` which further personalizes them per recipient.

---

### 6. Personalization at Every Level

Mass emails with zero personalization are easy for both humans and filters to identify. Real outreach references something specific about the person or their company.

**What we do:** Every email addresses the recipient by name, references their exact job title, and names their company — in the subject line, the opening paragraph, and the body. This is injected from your contacts Excel at send time.

---

### 7. Retry with Exponential Backoff

Retrying immediately after a failure often hits the same rate-limit wall. Exponential backoff waits progressively longer between attempts.

**What we do:** On SMTP failure, the script waits 2s before attempt 2, then 4s before attempt 3, then 8s before attempt 4 — and reconnects SMTP between attempts.

---

### 8. Duplicate & Format Validation

Sending to invalid or duplicate addresses increases your bounce rate. A high bounce rate damages sender reputation and leads to future emails being filtered.

**What we do:** Every address is validated for format using `email-validator` before sending. Duplicates within the same run are tracked in a `sent_emails` set and skipped.

---

## 📋 Logging

Three Excel log files are auto-generated (and gitignored):

| File | Contents |
|------|----------|
| `sent_emails.xlsx` | All successfully sent emails with name, company, subject |
| `failed_emails.xlsx` | Failed sends with the error reason |
| `skipped_emails.xlsx` | Duplicates and invalid addresses with reason |

Logs are saved **after every single email**, so if the script crashes mid-run, you don't lose your progress data.

---

## 🔁 Retry Logic

Each email gets up to `MAX_RETRIES` attempts (default: 3). On failure:

1. Catches both `SMTPServerDisconnected` and all `SMTPException` subclasses
2. Waits with exponential backoff (2s → 4s → 8s)
3. Reconnects SMTP fresh before retrying
4. Logs as failed after all attempts are exhausted

Unknown / unexpected errors are logged immediately without retry (no point retrying a `ValueError`).

---

## 🖨️ Console Output

While running, the script prints a live feed:

```
✅ SMTP connected.
Loaded 96 opening/subject variants from Opening_Lines.xlsx
Total usable contacts: 312

✅ Sent (1) → Riya Sharma <riya@zepto.com>
   Subject : Internship Inquiry – Mihir Kulkarni | AI/ML & Full-Stack
   Opening : Your profile as HR Manager at Zepto stood out to me...
   ⏳ Waiting 47s...

✅ Sent (2) → Arjun Mehta <arjun@razorpay.com>
   Subject : Seeking Internship at Razorpay – Mihir Kulkarni
   Opening : I came across your work as CTO at Razorpay and was genuinely inspired...
   ⏳ Waiting 62s...

⏸️  Batch of 20 done. Cooling down for 11 mins...

✅ SMTP connected.
...

========================================
   EMAIL CAMPAIGN COMPLETE
========================================
  ✅ Sent:    312
  ❌ Failed:  4
  ⚠️  Skipped: 11
========================================
```

---

## ❗ Important Notes

- **Gmail daily limit:** Free Gmail accounts can send ~500 emails/day. Stay under this.
- **App Password required:** Regular Gmail password won't work with SMTP. Generate an App Password as described in Setup.
- **Test first:** Always set `MAX_EMAILS_TO_SEND = 5` and run against your own email addresses before a real campaign.
- **Contacts file is gitignored:** Any file matching `*.xlsx` (except `sample.xlsx` and `Opening_Lines.xlsx`) is excluded from git. Never commit real contact data.
- **Resume filename:** Make sure `RESUME_FILE` in the script matches your actual PDF filename exactly.

---

## 📜 .gitignore

```
.env
__pycache__/
*.pyc
contacts.xlsx
contacts 2.xlsx
contacts 3.xlsx
contacts 4.xlsx
sent_emails.xlsx
failed_emails.xlsx
skipped_emails.xlsx
*.xlsx
```

> `sample.xlsx` and `Opening_Lines.xlsx` are committed because they don't contain personal data — they're part of the project structure.

---

## 👤 Author

**Mihir Kulkarni**
B.Tech Computer Engineering · RAIT, D. Y. Patil University · Navi Mumbai
[mih.kul.work@gmail.com](mailto:mih.kul.work@gmail.com) · [linkedin.com/in/mihir2005](https://linkedin.com/in/mihir2005) · [github.com/Mihirkool](https://github.com/Mihirkool)