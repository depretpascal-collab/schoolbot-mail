<!-- LOVABLE:BEGIN -->
> [!IMPORTANT]
> This project is connected to [Lovable](https://lovable.dev). Avoid rewriting
> published git history — force pushing, or rebasing/amending/squashing commits
> that are already pushed — as it rewrites history on Lovable's side and the
> user will likely lose their project history.
>
> Commits you push to the connected branch sync back to Lovable and show up in
> the editor, so keep the branch in a working state.
<!-- LOVABLE:END -->

## SchoolBot Mail (public/schoolbot-mail.py)

- The desktop app must stay dependency-free (Python standard library only) so PyInstaller can bundle it into a single double-clickable program.
- Mail content must never leave the user's computer: classification and drafting run against a local model (Ollama). Never route mail text through a cloud API by default — that would make each school use an invisible sub-processor under GDPR.
- Microsoft 365 mailboxes are reached with the device-code OAuth flow (app registration, no stored password) and Microsoft Graph; IMAP/SMTP with a password stays only as a manual fallback.
- The single Microsoft client id is a one-time Educlan setup step kept in `MS_CLIENT_ID`; never invent or hardcode a third party's client id.
