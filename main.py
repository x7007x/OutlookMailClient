import os, sys, json, re
from playwright.sync_api import sync_playwright

class OutlookClient:
    def __init__(self, email, password, headless=True):
        self.email = email
        self.password = password
        self.new_password = os.environ.get("OUTLOOK_NEW_PASSWORD", f"{password}@1")
        self.headless = headless
        self.alt_passwords = [p.strip() for p in os.environ.get("OUTLOOK_ALT_PASSWORDS", "").split(",") if p.strip()]
        self._pwd_candidates = [password] + self.alt_passwords
        self._pwd_idx = 0
        self._pwd_fails = 0

    def _visible(self, page, selector):
        try:
            return page.locator(selector).first.is_visible()
        except:
            return False

    def _text(self, page, selector):
        try:
            return (page.locator(selector).first.text_content() or "").strip()
        except:
            return ""

    def _wait_any(self, page, selectors, timeout=10000):
        combined = ", ".join(selectors)
        try:
            page.wait_for_selector(combined, timeout=timeout)
            return True
        except:
            return False

    def _login_flow(self, page):
        print("  → loading outlook.office.com/mail")
        page.goto("https://outlook.office.com/mail", wait_until="domcontentloaded")

        for i in range(40):
            url = page.url

            if "outlook.office.com/mail" in url and "/mail/" in url:
                print("  → inbox reached")
                return True

            has_email = self._visible(page, 'input[type="email"][name="loginfmt"]')
            has_pwd = self._visible(page, 'input[type="password"][name="passwd"]')

            if has_email and has_pwd:
                try:
                    filled = page.locator('input[type="email"][name="loginfmt"]').input_value()
                    if filled == self.email:
                        self._wait_any(page, [
                            "#idSubmit_ProofUp_Redirect",
                            "div[role='listbox']",
                            "button:has-text('Skip')",
                            "#currentPassword",
                        ], 3000)
                        continue
                except:
                    pass
                print("  → filling email + password (combined form)")
                page.locator('input[type="email"][name="loginfmt"]').fill(self.email)
                page.locator('input[type="password"][name="passwd"]').fill(self.password)
                page.locator('input[type="submit"]').first.click()
                self._wait_any(page, [
                    "#idSubmit_ProofUp_Redirect",
                    "div[role='listbox']",
                    "button:has-text('Skip')",
                    "#currentPassword",
                    "strong:has-text('Stay signed in')",
                    "h2[data-testid='reskin-step-title']",
                    "div[role='heading']",
                ], 10000)
                continue

            if has_email and not has_pwd:
                print("  → filling email")
                page.locator('input[type="email"][name="loginfmt"]').fill(self.email)
                page.locator('input[type="submit"]').first.click()
                self._wait_any(page, [
                    'input[type="password"][name="passwd"]',
                    "#idSubmit_ProofUp_Redirect",
                    "#currentPassword",
                ], 8000)
                continue

            if has_pwd and not has_email:
                if self._visible(page, 'div[role="alert"]'):
                    self._pwd_fails += 1
                    if self._pwd_fails >= len(self._pwd_candidates) * 2:
                        print("  ✗ all passwords exhausted")
                        sys.exit(1)
                    self._pwd_idx = (self._pwd_idx + 1) % len(self._pwd_candidates)
                    page.locator('input[type="password"][name="passwd"]').wait_for(state="visible", timeout=5000)
                    continue
                pw = self._pwd_candidates[self._pwd_idx % len(self._pwd_candidates)]
                print(f"  → filling password (candidate {self._pwd_idx % len(self._pwd_candidates) + 1})")
                page.locator('input[type="password"][name="passwd"]').fill(pw)
                page.locator('input[type="submit"]').first.click()
                self._wait_any(page, [
                    "#idSubmit_ProofUp_Redirect",
                    "div[role='listbox']",
                    "button:has-text('Skip')",
                    "#currentPassword",
                    "strong:has-text('Stay signed in')",
                    "div[role='heading']",
                ], 10000)
                continue

            if self._visible(page, "#currentPassword"):
                print("  → forced password change")
                page.locator("#currentPassword").fill(self.password)
                page.locator("#newPassword").fill(self.new_password)
                page.locator("#confirmNewPassword").fill(self.new_password)
                page.locator("#idSIButton9").click()
                self._wait_any(page, [
                    "#idSubmit_ProofUp_Redirect",
                    "div[role='listbox']",
                    'input[type="password"][name="passwd"]',
                ], 10000)
                continue

            if self._visible(page, "#idSubmit_ProofUp_Redirect"):
                print("  → MFA proof-up, redirecting")
                page.locator("#idSubmit_ProofUp_Redirect").click()
                self._wait_any(page, [
                    "button:has-text('Skip')",
                    "button:has-text('Later')",
                    "div[role='listbox']",
                ], 10000)
                continue

            if "mysignins.microsoft.com" in url:
                print("  → mysignins registration page")
                skipped = False
                for sel in ['button:has-text("Skip")', 'button:has-text("Later")', 'button:has-text("Not now")', 'a:has-text("Do this later")']:
                    try:
                        if page.locator(sel).is_visible():
                            page.locator(sel).click()
                            skipped = True
                            break
                    except:
                        continue
                if not skipped:
                    print("  → navigating directly to Outlook")
                    page.goto("https://outlook.office.com/mail", wait_until="domcontentloaded")
                self._wait_any(page, ["div[role='listbox']", "div[role='list']"], 10000)
                continue

            try:
                h2 = page.locator('h2[data-testid="reskin-step-title"]').first
                if h2.is_visible() and "Microsoft Authenticator" in (h2.text_content() or ""):
                    print("  → skipping Authenticator setup")
                    for sel in ['button:has-text("Skip setup")', 'button:has-text("Skip for now")']:
                        btn = page.locator(sel)
                        if btn.is_visible():
                            btn.click()
                            break
                    self._wait_any(page, ["div[role='listbox']"], 8000)
                    continue
            except:
                pass

            if "SAS/ProcessAuth" in url or "reprocess" in url:
                print("  → SAML / reprocess form")
                form = page.locator('form[action*="SAML"], form[action*="saml"], form[name="hiddenform"]')
                if form.is_visible():
                    form.first.evaluate("f => f.submit()")
                else:
                    page.locator('input[type="submit"]').first.click()
                self._wait_any(page, ["div[role='listbox']"], 10000)
                continue

            if self._visible(page, 'div[role="heading"]') and "Keep me signed in" in (page.locator('div[role="heading"]').first.text_content() or ""):
                print("  → KMSI: Yes")
                page.locator('input[type="submit"][value="Yes"]').click()
                self._wait_any(page, ["div[role='listbox']"], 8000)
                continue

            try:
                if page.locator('strong:has-text("Stay signed in")').is_visible():
                    print("  → Stay signed in: No")
                    page.locator('input[type="button"][value="No"]').click()
                    self._wait_any(page, ["div[role='listbox']"], 8000)
                    continue
            except:
                pass

            self._wait_any(page, [
                'input[type="email"]',
                'input[type="password"]',
                "#idSubmit_ProofUp_Redirect",
                "#currentPassword",
                "div[role='listbox']",
                "button:has-text('Skip')",
            ], 5000)

        return False

    def login(self):
        with sync_playwright() as pw:
            browser = pw.chromium.launch(headless=self.headless, args=["--no-sandbox", "--disable-dev-shm-usage"])
            context = browser.new_context(
                viewport={"width": 1280, "height": 900},
                user_agent="Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36",
            )
            page = context.new_page()
            page.set_default_timeout(25000)

            ok = self._login_flow(page)
            if not ok:
                print("  ✗ login failed")
                sys.exit(1)

            cookies = context.cookies()
            browser.close()

        return cookies

    def get_emails(self, cookies):
        with sync_playwright() as pw:
            browser = pw.chromium.launch(headless=self.headless, args=["--no-sandbox", "--disable-dev-shm-usage"])
            context = browser.new_context(
                viewport={"width": 1280, "height": 900},
                user_agent="Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36",
            )
            context.add_cookies(cookies)
            page = context.new_page()

            print("  → loading inbox with cookies")
            page.goto("https://outlook.office.com/mail", wait_until="domcontentloaded")
            self._wait_any(page, [
                'div[role="listbox"]',
                'div[role="list"]',
                '[data-convid]',
            ], 20000)

            emails = []
            seen = set()
            items = page.locator('[data-convid], [role="option"][tabindex], [role="listitem"]').all()
            for row in items:
                try:
                    sender = (row.locator('span[title]').first.text_content() or "").strip()
                    subject = (row.locator('span.TtcXM, span[data-automationid="subject"]').first.text_content() or "").strip()
                    preview = (row.locator('span.FqgPc, span[data-automationid="preview"]').first.text_content() or "").strip()
                    time = (row.locator('span._rWRU, span[data-automationid="time"]').first.text_content() or "").strip()
                    key = f"{sender}|{subject}"
                    if key and key not in seen:
                        seen.add(key)
                        emails.append({"sender": sender, "subject": subject, "preview": preview, "time": time})
                except:
                    continue

            if emails:
                print(f"  → reading details for {len(emails)} email(s)")
                first = page.locator('[data-convid], [role="option"][tabindex], [role="listitem"]').first
                if first.is_visible():
                    first.click()
                    self._wait_any(page, [
                        '[data-testid="message-body"]',
                        'div[aria-label="Message body"]',
                        'div[role="heading"][aria-level="3"]',
                    ], 10000)

                    for idx, email in enumerate(emails):
                        try:
                            data = {}
                            data["subject"] = self._text(page, 'div[role="heading"][aria-level="3"] span.JdFsz, [data-automationid="subjectLine"]')

                            from_el = page.locator('span[title][title*="@"]').first
                            if from_el.is_visible():
                                data["from"] = {"name": (from_el.text_content() or "").strip(), "email": (from_el.get_attribute("title") or "").strip()}
                            else:
                                data["from"] = {"name": "", "email": ""}

                            data["to"] = []
                            try:
                                label = page.locator('span:text-is("To:")').first
                                if label.is_visible():
                                    for span in page.locator('span:text-is("To:") ~ span').all():
                                        data["to"].append({"name": (span.text_content() or "").strip(), "email": (span.get_attribute("title") or "").strip()})
                            except:
                                pass

                            time_el = page.locator("time").first
                            if time_el.is_visible():
                                data["date"] = (time_el.text_content() or "").strip()
                            else:
                                full = self._text(page, 'div[data-app-section="MailReadCompose"]')
                                m = re.search(r'(Mon|Tue|Wed|Thu|Fri|Sat|Sun)\s+\d{1,2}/\d{1,2}/\d{4}\s+\d{1,2}:\d{2}\s*(AM|PM)', full)
                                data["date"] = m.group(0) if m else ""

                            body_html = ""
                            raw_text = ""
                            for sel in ['[data-testid="message-body"]', 'div[aria-label="Message body"]', 'div[class*="allowTextSelection"]', 'div[data-app-section="MailReadCompose"]']:
                                if self._visible(page, sel):
                                    try:
                                        body_html = page.locator(sel).first.inner_html()
                                    except:
                                        pass
                                    try:
                                        raw_text = page.locator(sel).first.inner_text()
                                    except:
                                        pass
                                    break

                            lines = []
                            skip_pfx = ("Reply", "Forward", "Summarize", "More actions", "Customize Actions", "Flag this", "Keep this")
                            for line in raw_text.splitlines():
                                s = line.strip()
                                if not s: continue
                                if any(s.startswith(p) for p in skip_pfx): continue
                                if len(s) == 1 and ord(s) > 127: continue
                                if "CAUTION:" in s: continue
                                lines.append(s)
                            data["body_text"] = "\n".join(lines).strip()
                            data["body_html"] = body_html

                            email["details"] = data
                        except:
                            email["details"] = None

                        if idx == 0:
                            break

            browser.close()
        return emails


if __name__ == "__main__":
    email = os.environ.get("OUTLOOK_EMAIL", "")
    password = os.environ.get("OUTLOOK_PASSWORD", "")

    client = OutlookClient(email, password)

    print("logging in...")
    cookies = client.login()
    print(f"  cookies: {len(cookies)}")

    print("fetching emails...")
    emails = client.get_emails(cookies)

    print(f"\nemails: {len(emails)}")
    print(json.dumps(emails, indent=2, ensure_ascii=False))
