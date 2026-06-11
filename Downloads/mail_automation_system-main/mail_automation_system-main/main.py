import pandas as pd
import openpyxl
import smtplib
import time
import os
import random
import itertools
from dotenv import load_dotenv
from email_validator import validate_email, EmailNotValidError
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication


# ==============================
# SETTINGS
# ==============================

CONTACTS_FILE      = r"..\contacts.xlsx"
RESUME_FILE        = r"..\hamsini_kuppam_Resume.pdf"
VARIANTS_FILE      = r"..\Opening_Lines.xlsx"

MIN_DELAY          = 35
MAX_DELAY          = 75
BATCH_SIZE         = 20
BATCH_BREAK_MIN    = 8 * 60
BATCH_BREAK_MAX    = 15 * 60
MAX_EMAILS_TO_SEND = 500
MAX_RETRIES        = 3


# ==============================
# LOAD ENVIRONMENT VARIABLES
# ==============================

load_dotenv()

EMAIL_ADDRESS  = os.getenv("EMAIL_ADDRESS")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")

if not EMAIL_ADDRESS or not EMAIL_PASSWORD:
    raise ValueError("EMAIL_ADDRESS or EMAIL_PASSWORD missing in .env file")


# ==============================
# LOAD OPENING LINES & SUBJECTS FROM EXCEL
# ==============================

def load_variants(filepath):
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"Variants file not found: {filepath}")

    wb = openpyxl.load_workbook(filepath)
    ws = wb.active

    rows = [
        (row[0], row[1])
        for row in ws.iter_rows(min_row=2, values_only=True)
        if row[0] and row[1]
    ]

    if not rows:
        raise ValueError("Opening_Lines.xlsx has no data rows.")

    print(f"Loaded {len(rows)} opening/subject variants from {filepath}")
    print("These will cycle infinitely — email 97 reuses row 1, email 98 reuses row 2, etc.")

    return itertools.cycle(rows)
import os

print("Current Working Directory:", os.getcwd())
print("Files in Directory:")
print(os.listdir())

variants_cycle = load_variants(VARIANTS_FILE)


# ==============================
# SMTP CONNECT
# ==============================

def connect_smtp():
    server = smtplib.SMTP("smtp.gmail.com", 587, timeout=60)
    server.ehlo()
    server.starttls()
    server.ehlo()
    server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
    print("✅ SMTP connected.")
    return server


# ==============================
# EMAIL VALIDATION
# ==============================

def is_valid_email_format(email):
    try:
        validate_email(email, check_deliverability=False)
        return True
    except EmailNotValidError:
        return False


# ==============================
# PLAIN TEXT BODY
# ==============================

def build_plain_body(name, designation, company, opening_line):
    resolved_opening = opening_line.format(
        name=name, designation=designation, company=company
    )
    return (
        f"Hello {name},\n\n"
        f"{resolved_opening}\n\n"
        f"My focus areas: LLMs & Agentic AI, Full-Stack Development, Data Analytics, UI/UX Design, Power BI, SQL.\n\n"
        f"Recent work:\n"
        f"- Agentic Financial News Pipeline (LLMs, RAG, Supabase)\n"
        f"- Miraki: Full-stack digital art marketplace (React, MongoDB, Node.js, Vercel)\n\n"
        f"Also serving as Creative Head at CSI-RAIT — organized workshops for 150-200+ students.\n\n"
        f"Resume is attached. I'd love to connect — even a 10-minute call would mean a lot.\n\n"
        f"Thank you for your time.\n\n"
        f"Warm regards,\n"
        f"Hamsini Kuppam\n"
        f"hamsini.kuppam@gmail.com\n"
        f"linkedin.com/in/hamsini-kuppam\n"
        f"github.com/Hamsini2606"
    )


# ==============================
# HTML EMAIL BODY
# ==============================

def build_html_body(name, designation, company, opening_line):
    resolved_opening = opening_line.format(
        name=name, designation=designation, company=company
    )

    template = """\
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
</head>

<body style="margin:0;padding:0;background-color:#f0f2f5;font-family:'Segoe UI',Arial,sans-serif;">
  <table width="100%" cellpadding="0" cellspacing="0" style="background-color:#f0f2f5;padding:32px 16px;">
    <tr>
      <td align="center">

        <table width="620" cellpadding="0" cellspacing="0"
          style="background-color:#ffffff;border-radius:10px;overflow:hidden;box-shadow:0 2px 12px rgba(0,0,0,0.08);">

          <tr>
            <td style="background:linear-gradient(90deg,#1a1a2e 0%,#16213e 50%,#0f3460 100%);height:5px;"></td>
          </tr>

          <tr>
            <td style="padding:36px 40px 24px 40px;border-bottom:1px solid #eef0f4;">
              <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                  <td>
                    <p style="margin:0 0 4px 0;font-size:11px;font-weight:600;letter-spacing:1.5px;text-transform:uppercase;color:#0f3460;">
                      Internship Inquiry
                    </p>
                    <h1 style="margin:0;font-size:22px;font-weight:700;color:#1a1a2e;">
                      Hamsini Kuppam
                    </h1>
                    <p style="margin:6px 0 0 0;font-size:13px;color:#6b7280;">
                      Computer Engineering &middot; Artificial Intelligence &middot; Data Science &middot; Full-Stack Development
                    </p>
                  </td>
                  <td align="right" valign="top">
                    <div style="display:inline-block;background-color:#eef4ff;border:1px solid #c7d9ff;border-radius:20px;padding:5px 14px;">
                      <span style="font-size:12px;font-weight:600;color:#2563eb;">
                        CGPA 9.84 / 10
                      </span>
                    </div>
                  </td>
                </tr>
              </table>
            </td>
          </tr>

          <tr>
            <td style="padding:32px 40px;">

              <p style="margin:0 0 18px 0;font-size:15px;color:#374151;line-height:1.7;">
                Hello <strong style="color:#1a1a2e;">{name}</strong>,
              </p>

              <p style="margin:0 0 18px 0;font-size:15px;color:#374151;line-height:1.7;">
                {resolved_opening}
              </p>

              <p style="margin:0 0 24px 0;font-size:15px;color:#374151;line-height:1.7;">
                My interests span <strong>Artificial Intelligence</strong>,
                <strong>Machine Learning</strong>,
                <strong>Data Analytics</strong>,
                and <strong>Full-Stack Development</strong>, and I actively apply them through projects and internships:
              </p>

              <table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:24px;">
                <tr>
                  <td width="48%" valign="top"
                    style="padding:14px 16px;background-color:#f8fafc;border-radius:8px;border-left:3px solid #0f3460;">
                    <p style="margin:0 0 4px 0;font-size:11px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:#6b7280;">
                      AI / Machine Learning
                    </p>
                    <p style="margin:0;font-size:13.5px;color:#1a1a2e;line-height:1.55;">
                      Built <strong>Maargadarshak &ndash; AI Career Guidance Platform</strong>
                      using Python, React, Machine Learning, and REST APIs to provide
                      personalized career recommendations based on student profiles and skills.
                    </p>
                  </td>
                  <td width="4%"></td>
                  <td width="48%" valign="top"
                    style="padding:14px 16px;background-color:#f8fafc;border-radius:8px;border-left:3px solid #0f3460;">
                    <p style="margin:0 0 4px 0;font-size:11px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:#6b7280;">
                      Full-Stack Development
                    </p>
                    <p style="margin:0;font-size:13.5px;color:#1a1a2e;line-height:1.55;">
                      Developed <strong>Miraki &ndash; Artistry Hub</strong>,
                      a full-stack collaboration platform for artists and clients using
                      React, Tailwind CSS, and REST APIs with responsive user interfaces
                      and portfolio showcase features.
                    </p>
                  </td>
                </tr>
              </table>

              <p style="margin:0 0 18px 0;font-size:15px;color:#374151;line-height:1.7;">
                I have completed internships as a
                <strong>Data Science Intern at ShadowFox</strong>,
                <strong>Artificial Intelligence Intern at CodSoft</strong>,
                <strong>Salesforce Developer Intern through Salesforce SmartBridge</strong>,
                and <strong>Data Analyst Intern at Edunet Foundation (TechSaksham &ndash; SAP &amp; Microsoft)</strong>.
                These experiences strengthened my skills in Python, SQL, Power BI,
                Salesforce, CRM Development, Data Analytics, and AI-based solutions.
              </p>

              <p style="margin:0 0 28px 0;font-size:15px;color:#374151;line-height:1.7;">
                I've attached my resume for your reference. If my profile aligns with any
                opportunities on your team, I would be grateful for the opportunity to connect
                and discuss how I can contribute.
              </p>

              <table cellpadding="0" cellspacing="0" style="margin-bottom:32px;">
                <tr>
                  <td style="background-color:#0f3460;border-radius:6px;padding:12px 24px;">
                    <a href="mailto:khamsini2606@gmail.com"
                      style="color:#ffffff;font-size:14px;font-weight:600;text-decoration:none;">
                      View Resume &amp; Connect &rarr;
                    </a>
                  </td>
                </tr>
              </table>

              <p style="margin:0;font-size:15px;color:#374151;line-height:1.7;">
                Thank you for your time and consideration. I look forward to hearing from you.
              </p>

            </td>
          </tr>

          <tr>
            <td style="padding:24px 40px 32px 40px;border-top:1px solid #eef0f4;background-color:#fafbfc;">
              <p style="margin:0 0 2px 0;font-size:14px;font-weight:700;color:#1a1a2e;">
                Hamsini Kuppam
              </p>
              <p style="margin:0 0 10px 0;font-size:12.5px;color:#6b7280;">
                B.Tech Computer Engineering &middot; Ramrao Adik Institute of Technology (RAIT), Navi Mumbai
              </p>
              <table cellpadding="0" cellspacing="0">
                <tr>
                  <td style="padding-right:16px;">
                    <a href="mailto:khamsini2606@gmail.com"
                      style="font-size:12.5px;color:#2563eb;text-decoration:none;">
                      khamsini2606@gmail.com
                    </a>
                  </td>
                  <td style="padding-right:16px;">
                    <a href="https://linkedin.com/in/hamsini-kuppam"
                      style="font-size:12.5px;color:#2563eb;text-decoration:none;">
                      LinkedIn
                    </a>
                  </td>
                  <td>
                    <a href="https://github.com/hamsini2606"
                      style="font-size:12.5px;color:#2563eb;text-decoration:none;">
                      GitHub
                    </a>
                  </td>
                </tr>
              </table>
            </td>
          </tr>

          <tr>
            <td style="background:linear-gradient(90deg,#1a1a2e 0%,#0f3460 100%);height:4px;"></td>
          </tr>

        </table>
      </td>
    </tr>
  </table>
</body>
</html>"""

    return template.format(name=name, resolved_opening=resolved_opening)


# ==============================
# LOAD CONTACTS
# ==============================

df = pd.read_excel(CONTACTS_FILE)
df.columns = df.columns.astype(str).str.strip().str.lower()

print("Excel columns found:", df.columns.tolist())

required_columns = ["name", "designation", "company", "mail"]
for col in required_columns:
    if col not in df.columns:
        raise ValueError(f"Missing column in Excel: '{col}'")

df = df.dropna(subset=required_columns)
for col in required_columns:
    df[col] = df[col].astype(str).str.strip()

df = df[
    (df["name"] != "") & (df["designation"] != "") &
    (df["company"] != "") & (df["mail"] != "")
]

print(f"Total usable contacts: {len(df)}")


# ==============================
# LOAD RESUME
# ==============================

if not os.path.exists(RESUME_FILE):
    raise FileNotFoundError(f"Resume not found: {RESUME_FILE}")

with open(RESUME_FILE, "rb") as f:
    resume_data = f.read()

print(f"Resume loaded: {RESUME_FILE} ({len(resume_data) // 1024} KB)")


# ==============================
# LOG SETUP
# ==============================

sent_logs    = []
failed_logs  = []
skipped_logs = []
sent_emails  = set()


def save_logs():
    if sent_logs:
        pd.DataFrame(sent_logs).to_excel("sent_emails.xlsx", index=False)
    if failed_logs:
        pd.DataFrame(failed_logs).to_excel("failed_emails.xlsx", index=False)
    if skipped_logs:
        pd.DataFrame(skipped_logs).to_excel("skipped_emails.xlsx", index=False)


# ==============================
# SEND EMAILS
# ==============================

server = connect_smtp()
emails_sent_count = 0

for index, row in df.iterrows():

    if emails_sent_count >= MAX_EMAILS_TO_SEND:
        print(f"🛑 Batch limit reached: {MAX_EMAILS_TO_SEND} emails sent.")
        break

    name        = row["name"]
    designation = row["designation"]
    company     = row["company"]
    email       = row["mail"].lower().strip()

    # Skip duplicates
    if email in sent_emails:
        print(f"⚠️  Skipping duplicate: {email}")
        skipped_logs.append({"name": name, "designation": designation, "company": company, "mail": email, "reason": "Duplicate"})
        save_logs()
        continue

    sent_emails.add(email)

    # Skip invalid format
    if not is_valid_email_format(email):
        print(f"⚠️  Skipping invalid format: {email}")
        skipped_logs.append({"name": name, "designation": designation, "company": company, "mail": email, "reason": "Invalid email format"})
        save_logs()
        continue

    # Pull NEXT variant from the cycle
    opening_line, subject_template = next(variants_cycle)

    # Resolve {company} in subject if present
    subject = subject_template.format(
        name=name, designation=designation, company=company
    )

    # Build both parts of the email
    html_body  = build_html_body(name, designation, company, opening_line)
    plain_body = build_plain_body(name, designation, company, opening_line)

    # Multipart/alternative (HTML + plain text = better deliverability)
    alt = MIMEMultipart("alternative")
    alt.attach(MIMEText(plain_body, "plain"))
    alt.attach(MIMEText(html_body,  "html"))

    # Wrap in mixed so we can add the PDF attachment
    outer              = MIMEMultipart("mixed")
    outer["From"]      = EMAIL_ADDRESS
    outer["To"]        = email
    outer["Subject"]   = subject
    outer.attach(alt)

    resume_attachment = MIMEApplication(resume_data, _subtype="pdf")
    resume_attachment.add_header("Content-Disposition", "attachment", filename=RESUME_FILE)
    outer.attach(resume_attachment)

    # Send with retry + exponential backoff
    success = False

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            failed_recipients = server.sendmail(EMAIL_ADDRESS, email, outer.as_string())

            if failed_recipients:
                print(f"❌ Rejected: {email} &rarr; {failed_recipients}")
                failed_logs.append({"name": name, "designation": designation, "company": company, "mail": email, "reason": str(failed_recipients)})
            else:
                print(f"✅ Sent ({emails_sent_count + 1}) &rarr; {name} <{email}>")
                print(f"   Subject : {subject[:60]}")
                print(f"   Opening : {opening_line[:70]}...")
                sent_logs.append({
                    "name": name, "designation": designation,
                    "company": company, "mail": email,
                    "subject": subject, "status": "Sent"
                })
                emails_sent_count += 1
                success = True

            break

        except (smtplib.SMTPServerDisconnected, smtplib.SMTPException) as e:
            print(f"⚠️  SMTP error attempt {attempt}/{MAX_RETRIES} for {email}: {e}")
            if attempt < MAX_RETRIES:
                wait = 2 ** attempt
                print(f"   Reconnecting in {wait}s...")
                time.sleep(wait)
                try:
                    server = connect_smtp()
                except Exception as re:
                    print(f"   Reconnect failed: {re}")
            else:
                print(f"❌ Giving up on {email}.")
                failed_logs.append({"name": name, "designation": designation, "company": company, "mail": email, "reason": str(e)})

        except Exception as e:
            print(f"❌ Unexpected error for {email}: {e}")
            failed_logs.append({"name": name, "designation": designation, "company": company, "mail": email, "reason": str(e)})
            break

    save_logs()

    # Batch cooling break every BATCH_SIZE emails
    if success and emails_sent_count % BATCH_SIZE == 0:
        break_time = random.randint(BATCH_BREAK_MIN, BATCH_BREAK_MAX)
        print(f"\n⏸️  Batch of {BATCH_SIZE} done. Cooling down for {break_time // 60} mins...\n")
        time.sleep(break_time)
        try:
            server.quit()
        except Exception:
            pass
        server = connect_smtp()

    elif success:
        delay = random.randint(MIN_DELAY, MAX_DELAY)
        print(f"   ⏳ Waiting {delay}s...\n")
        time.sleep(delay)
    else:
        time.sleep(3)


# ==============================
# CLOSE & SUMMARISE
# ==============================

try:
    server.quit()
    print("SMTP connection closed.")
except Exception as e:
    print(f"SMTP quit error (non-critical): {e}")

save_logs()

print("\n" + "=" * 40)
print("   EMAIL CAMPAIGN COMPLETE")
print("=" * 40)
print(f"  ✅ Sent:    {len(sent_logs)}")
print(f"  ❌ Failed:  {len(failed_logs)}")
print(f"  ⚠️  Skipped: {len(skipped_logs)}")
print("=" * 40)