# 📧 Automated Personalized Email & Invitation Sender

## 📖 Background Story

I had the responsibility of sending **personalized invitation emails and custom invitation cards** to each faculty member of my university to invite them to **GDSC TechXpo**.

Manually sending emails, attaching the correct invitation card for each faculty member, and ensuring personalization would have been extremely time-consuming and error-prone.

As the saying goes, *“Where there is a problem, there is always a way to build a solution.”*  
So, I approached this challenge with a **problem-solving mindset** and built a **Python automation script** that sends personalized emails and invitation cards in one go.

If you ever face a similar situation, this workflow can save you **hours of manual effort**.

---

## 🚀 What This Script Does

- Reads recipient details from an Excel file  
- Sends HTML-formatted personalized emails  
- Attaches a unique invitation card per recipient  
- Supports CC recipients  
- Uses SMTP throttling to avoid rate limits  

---

## 📂 Project Structure

```
project/
│
├── send_emails.py          # Main Python script
├── NamelistPhase1.xlsx     # Excel file with recipient data
├── phase1/                 # Folder containing invitation cards
│   ├── Faculty1.png
│   ├── Faculty2.png
│   └── ...
└── README.md
```

---

## 📊 Excel File Schema (Required)

Your Excel file **must follow this schema exactly**:

| Column Name  | Description |
|-------------|------------|
| Email       | Recipient’s email address |
| EmailName   | Name used in the email greeting |
| Filepath    | Exact filename of the invitation card |

### Example Row

| Email | EmailName | Filepath |
|------|-----------|----------|
| john.doe@university.edu | Dr. Doe | John_Doe_Invite.png |

---

## 🧠 Excel Tips (Highly Recommended)

### Extracting Last Name Automatically

If you store full names and want to extract the last name:

```excel
=TRIM(RIGHT(A2, LEN(A2) - FIND("§", SUBSTITUTE(A2, " ", "§", LEN(A2)-LEN(SUBSTITUTE(A2," ",""))))))
```

This helps maintain polite and consistent greetings.

### Auto-Generating File Names

To automatically generate invitation filenames:

```excel
=SUBSTITUTE(A2," ","_") & "_Invite.png"
```

This reduces human error and ensures filename consistency.

---

## 🎨 Creating Invitation Cards in Canva

### Canva Bulk Create Workflow

1. Design **one invitation template** in Canva  
2. Add text placeholders like:
   - `{{Name}}`
   - `{{Department}}`
3. Use **Canva → Bulk Create**
4. Upload the Excel/CSV file
5. Map Excel columns to placeholders
6. Export all designs as **PNG**
7. Place all exported files into the `phase1/` folder

⚠️ The exported filenames **must match** the `Filepath` column exactly.

---

## ⚙️ Script Configuration

Update these values in the script:

```python
EXCEL_PATH = "path_to_excel.xlsx"
ATTACHMENT_BASE_PATH = "path_to_invitation_cards"
EMAIL = "your_email@domain.edu"
PASSWORD = "your_app_password"
CC_EMAIL = "cc_email@domain.edu"
```

🔐 **Security Note:**  
Never commit real credentials to GitHub. Use environment variables for production.

---

## 🔁 How the Script Works

1. Loads recipient data from Excel  
2. Connects to Outlook SMTP (Office365)  
3. Builds a personalized HTML email  
4. Attaches the correct invitation card  
5. Sends the email  
6. Waits briefly to avoid rate limits  
7. Repeats for all recipients  

---

## ⏱️ Rate Limiting

The script includes:

```python
time.sleep(2)
```

This prevents SMTP blocking and improves deliverability.

---

## ❌ Common Mistakes to Avoid

- File name mismatch between Excel and folder  
- Missing attachment files  
- Committing real passwords  
- Sending emails too quickly  

---

## 📌 Possible Use Cases

- Event invitations  
- Conference passes  
- Certificates distribution  
- Offer letters  
- Personalized announcements  

---

## 🤝 Final Thoughts

This project demonstrates how **automation, structured data, and creativity** can solve real-world organizational problems efficiently.

If you are sending hundreds of personalized emails — **don’t do it manually. Automate it.**
