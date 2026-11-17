# outlook-mcp: Outlook mail and calendar MCP server

An MCP server that lets Claude (or any MCP client) work with your Exchange
Online / Outlook mailbox as you: find and read mail, follow conversations,
read attachments, prepare drafts, keep the mailbox organised and, once you
approve, send mail on your behalf. It also reads your calendar, finds free
time (yours and your colleagues'), creates appointments and meetings, and
answers invitations.

It is the mail sibling of [msteams-mcp](https://github.com/dataoracle/msteams-mcp)
and shares its design. It uses **delegated** Microsoft Graph permissions: you
sign in once on the server (OAuth2 authorization code + PKCE), and the server
then keeps itself signed in with the refresh token, getting short-lived access
tokens as needed. Everything it does appears in Outlook as done by you.

It runs as a hardened systemd service. No credential is stored in plain text:

| Secret | Where it lives |
|---|---|
| Entra app client secret | TPM-sealed systemd credential |
| MCP access token | TPM-sealed systemd credential |
| Vault key (AES-256) | TPM-sealed systemd credential |
| Microsoft refresh token | `/var/lib/outlook-mcp/vault.json.enc`, AES-256-GCM encrypted with the vault key |
| Access tokens (1 h) | Process memory only |

**Why a vault?** Entra ID *rotates* refresh tokens: every refresh returns a new
one, and each one expires after about 90 days. The server has to save each new
token while it runs, but the sandboxed service can't use the TPM directly. So
the TPM seals a key, and that key encrypts the token file. A copy of the disk
or a backup without this machine's TPM can't read the token.

## Tools

| Tool | Returns / does |
|---|---|
| `whoami()` | The signed-in mailbox, sign-in state, inbox counts, and what the server may do (drafts, sending, approval mode, allowed domains, signature) |
| `list_folders()` | All mail folders, nested, with unread/total counts and ids |
| `list_messages(folder?, limit?, since?, until?, unread_only?, sender?, flagged_only?, has_attachments?, importance?, focused?, preview_chars?)` | One folder, newest first: time, sender (recipients in Sent/Drafts), subject, marks, preview, message id. `until` pages back |
| `search_messages(query?, sender?, to?, subject?, since?, until?, has_attachments?, folder?, to_me?, unread_only?, limit?)` | Mailbox search (KQL): keywords in subject/body/attachments, by sender or recipient name/address, by subject. `to_me` keeps only mail where you are in To/Cc |
| `get_message(message_id, body?, max_chars?, offset?, mark_read?)` | One message in full: headers, attachments with ids, conversation id, Outlook link, body as text. `body="new"` drops the quoted history |
| `get_thread(message_id \| conversation_id, limit?, max_chars?)` | The whole conversation across folders (including your replies), oldest first, each showing only its new text |
| `get_attachment(message_id, attachment_id, max_chars?, offset?)` | An attachment's text: PDF, Word, Excel, PowerPoint, text/CSV/JSON/HTML, attached emails |
| `find_people(query)` | A person's email by name: people you work with, the directory, and addresses from your own mail |
| `create_draft(to, subject, body, cc?, bcc?, format?, importance?, attachments?)` ✏️ | Save a new mail in Drafts |
| `create_reply_draft(message_id, body, reply_all?, format?, cc?, attachments?)` ✏️ | Reply draft; your text above the quoted original |
| `create_forward_draft(message_id, to, body?, cc?, format?, attachments?)` ✏️ | Forward draft, original attachments included |
| `update_draft(draft_id, subject?, body?, to?, cc?, bcc?, format?, importance?, attachments?)` ✏️ | Change a draft |
| `delete_draft(draft_id)` ✏️ | Delete a draft (drafts only) |
| `mark_read(message_ids, read?)` ✏️ | Mark read / unread (up to 100 at once) |
| `flag_messages(message_ids, status?, due?)` ✏️ | Follow-up flag: flagged (with due date), complete, clear |
| `set_categories(message_ids, categories, mode?)` ✏️ | Add / remove / replace Outlook categories |
| `move_messages(message_ids, folder)` ✏️ | Move to archive, Deleted Items (recoverable), junk or any folder |
| `create_folder(name, parent?)` ✏️ | New mail folder |
| `send_draft(draft_id)` ✉️ | Send an existing draft |
| `send_mail(to, subject, body, cc?, bcc?, format?, importance?, attachments?)` ✉️ | Send a new mail now (copy kept in Sent Items) |
| `reply(message_id, body, reply_all?, format?, cc?, attachments?)` ✉️ | Reply and send now |
| `forward(message_id, to, body?, cc?, format?, attachments?)` ✉️ | Forward and send now |

**Calendar** (registered when `ENABLE_CALENDAR_TOOLS=true`, the default):

| Tool | Returns / does |
|---|---|
| `list_calendars()` | Your calendars with ids, owner and whether you can edit them |
| `list_events(start?, end?, days?, calendar?, query?, needs_response?, limit?)` | Events in a range (default: today + 7 days), recurring meetings expanded: time, subject, place, Teams, your response, organiser, attendee replies, ids. `needs_response=True` lists unanswered invitations |
| `get_event(event_id)` | One event in full: time, recurrence, Teams link, organiser, every attendee's response, description |
| `find_free_time(start?, end?, duration_minutes?, attendees?, working_hours_only?, tentative_is_busy?)` | Open slots for you and, optionally, colleagues (their Outlook free/busy), within your working hours |
| `create_event(subject, start, end? \| duration_minutes?, all_day?, attendees?, optional_attendees?, location?, body?, online_meeting?, show_as?, reminder_minutes?, categories?, sensitivity?, calendar?, repeat?, repeat_until?, repeat_count?)` ✏️📅 | An appointment, or a meeting (invitations sent) when attendees are given; Teams link with `online_meeting=True`; daily/weekdays/weekly/biweekly/monthly series |
| `update_event(event_id, subject?, start?, end?, location?, body?, add_attendees?, add_optional_attendees?, remove_attendees?, online_meeting?, show_as?, reminder_minutes?, categories?)` ✏️📅 | Change an event; for a meeting you organise, attendees get an update |
| `respond_to_event(response, event_id \| message_id, comment?, send_response?, propose_start?, propose_end?)` ✏️📅 | Accept, tentatively accept or decline an invitation (from the calendar or the invitation email), optionally proposing a new time |
| `cancel_event(event_id, comment?)` ✏️📅 | Cancel a meeting you organise (attendees are notified), or delete an appointment |

Calendar times are shown and entered in your time zone (`TIMEZONE`, or your
Outlook setting), not UTC: `start="2026-11-18T09:30"` means 09:30 your time.
Dates (`2026-11-18`), `today`, `tomorrow` and `now` work too.

✏️ = write tools, registered only when `ENABLE_WRITE_TOOLS=true` (the default).
✉️ = send tools, registered only when `ENABLE_SEND_TOOLS=true` (the default).
With `ENABLE_SEND_TOOLS=false` the agent can prepare drafts but never send;
you send them from Outlook.
📅 = can notify other people (invitations, updates, responses, cancellations),
which is treated exactly like sending mail: see the approval section. With
`ENABLE_SEND_TOOLS=false` these tools still work on your own calendar, but
refuse anything that would notify someone.

**Writing.** Bodies take `format="text"` (as typed, line breaks kept),
`"markdown"` (bold, lists, links, tables) or `"html"`. Recipients are email
addresses, as a list or a `;`/`,`-separated string; `Name <a@b.com>` and
`"Last, First" <a@b.com>` are accepted, plain names are refused so the agent
can't guess a recipient. Small attachments (3 MB in total) can be added inline
as `[{name, content | content_base64, content_type?}]`. If `SIGNATURE_FILE` is
set, the signature is appended to every new mail, reply and forward (Graph
doesn't apply your Outlook signature).

**Output.** Read tools return compact text to keep the agent's token use low:
one block per message with a date line per day, times in UTC, Defender Safe
Links unwrapped to the real URL, and inline-image references shortened to
`[image]`. Every read tool also takes `output="json"`. For example,
`list_messages` returns:

```
3 messages in Inbox, newest first (times UTC)
2026-11-13
10:01 Alice Johnson <alice@contoso.com> · Q4 numbers · UNREAD !high 📎
  Hi, Can you check the Q4 numbers? Report: https://example.com/report?q=1 [image] Thanks
  id AAkALgAAAAAAHYQDEapmEc2byACqAC-EWg0A...
09:00 Bob Smith <bob@partner.com> · Vendor call
  The vendor call moved to Thursday.
  id AAkALgAAAAAAHYQDEapmEc2byACqAC-EWg0B...
```

Message ids are Graph *immutable* ids: they stay valid when a message is moved
to another folder, or when a draft is sent.

## Sending on your behalf: approval

Mail can't be recalled, so sending has three independent safeguards:

1. **The agent asks you.** The server instructions and the agent skill tell
   the agent to show you recipients, subject and full text and wait for your
   explicit approval before any send tool. Send tools are also annotated as
   destructive, so MCP clients that prompt for tool use (Claude Code, Claude
   Desktop) ask you before each call. **Don't add the send tools to an
   auto-approve list.**
2. **`SEND_CONFIRMATION=elicit`** makes the *server* ask you, independently of
   the agent: before each send it shows a prompt in your MCP client with the
   recipients, subject and text, and sends only if you choose "send". The
   approval is bound to that exact content: if the message changes in between,
   you're asked again. If the client can't show the prompt, nothing is sent.
   Use this if you want a hard guarantee; it needs a client with MCP
   elicitation support.
3. **`SEND_ALLOWED_DOMAINS`** (e.g. `contoso.com,partner.com`) refuses to send to
   any other domain. Drafts are still allowed; a refused reply or forward stays in
   Drafts.

**Calendar actions that notify people count as sending** and go through the
same safeguards: `create_event` with attendees (invitations),
`update_event` on a meeting you organise (updates), `respond_to_event` with
`send_response=True` (the default) and `cancel_event` on a meeting. With
`SEND_CONFIRMATION=elicit` you approve each one in a prompt showing the
subject, time, attendees and text. `SEND_ALLOWED_DOMAINS` applies to attendees,
and to the organiser when your response carries a comment or a proposed time
(a plain accept/decline to an external organiser is allowed). Appointments
without attendees, and changes to your own reminder, "show as" or categories,
never notify anyone. Add the 📅 tools to the client's ask list too.

## Requirements

- **Server:** Linux with systemd 250 or newer and a TPM 2.0 chip (for example
  Ubuntu 24.04 or Debian 12), plus Python 3.10+ with `venv` support.
- **Entra ID app registration** (App ID + tenant ID):
  - Delegated permissions with admin consent:
    - required: `openid`, `offline_access`, `User.Read`, `Mail.ReadWrite`, `Mail.Send`;
    - for the calendar tools: `Calendars.ReadWrite` (your calendars, events,
      invitations, and colleagues' free/busy for `find_free_time`);
    - recommended: `People.Read` and `User.ReadBasic.All` (`find_people` searches
      the directory and your frequent contacts; without them it only finds
      addresses that appear in your own mail), `MailboxSettings.Read` (your
      Outlook time zone for calendar times unless `TIMEZONE` is set, and
      `whoami` shows whether auto-reply is on).

    The server requests `https://graph.microsoft.com/.default`, so it gets
    exactly what has been consented; optional features degrade gracefully when a
    permission is missing. Without `Mail.Send`, drafts still work and the send
    tools return 403.
  - **Authentication → Add a platform → Web**, with the redirect URI
    `http://localhost`. You can use another localhost URI such as
    `http://localhost:8400/callback`; set `MS_REDIRECT_URI` to the exact same
    value. Nothing needs to listen on it: you copy the code from the browser's
    address bar.
  - **Certificates & secrets → New client secret.** Copy its *Value* (not the
    Secret ID) straight into `seal-credential.sh client_secret`. Note the expiry
    date: when the secret expires the server stops working until you seal a new
    one (see below).

You can reuse the msteams-mcp app registration by adding the mail permissions
to it, or create a separate one. Each server keeps its own sign-in.

## Installation

Run these on the server from a clone of this repository.

```bash
# 1. Install code, venv, service user, config file, vault key and systemd unit
sudo ./deploy/install.sh

# 2. Set your tenant and app id (and the sending options, see Configuration)
sudo nano /etc/outlook-mcp/outlook-mcp.env

# 3. Seal the secrets (input is read without echo and never written to disk;
#    vault_key was already generated by install.sh)
sudo /opt/outlook-mcp/bin/seal-credential.sh client_secret          # paste the app's secret Value
sudo /opt/outlook-mcp/bin/seal-credential.sh mcp_token --generate

# 4. Sign in to Microsoft (see below)
sudo outlook-mcp-authorize

# 5. Start the service and enable it at boot
sudo systemctl enable --now outlook-mcp

# 6. Check it
systemctl status outlook-mcp
journalctl -u outlook-mcp -n 30
curl -s localhost:8767/health                 # {"status":"ok","auth":"ok"}
```

Show the generated MCP token once, to put in your MCP client config:

```bash
sudo systemd-creds decrypt --name=mcp_token /etc/credstore.encrypted/outlook-mcp.mcp_token.cred -; echo
```

### Signing in

`sudo outlook-mcp-authorize`:

1. prints a `login.microsoftonline.com` URL. Open it in a browser on any
   machine and sign in with the account whose mailbox the server should use;
2. the browser is redirected to `http://localhost/?code=…&state=…` and shows
   "can't connect", which is expected;
3. you copy that full URL from the address bar and paste it at the prompt
   (input is hidden). The code is single-use, expires in about 10 minutes, and
   can only be redeemed with both the client secret and the PKCE verifier held
   by the script;
4. the script decrypts the client secret into memory, exchanges the code for
   tokens, checks `/me` and the inbox, and writes the refresh token to the
   encrypted vault, owned by the service user.

The running service picks up the new vault on its next request, so you don't
need to restart it.

### When you have to sign in again

The server keeps its sign-in alive by itself. Every refresh rotates the stored
refresh token, and a background refresh every `KEEPALIVE_HOURS` (12 by
default) keeps it fresh even when nobody is using the server. You only need to
run `sudo outlook-mcp-authorize` again when Entra ID rejects the token, for
example:

- your password was changed or reset, or your sessions were revoked;
- an admin revoked consent or disabled the app;
- a Conditional Access sign-in frequency policy, or an MFA requirement, demands
  a fresh interactive sign-in;
- the server was off for longer than the refresh token's lifetime (about 90 days);
- you added permissions to the app (sign in again so the new consent is used).

When that happens, every tool returns *"The saved Microsoft sign-in is no
longer valid (AADSTS…). Run on the server: sudo outlook-mcp-authorize"*,
`/health` reports `"auth":"reauth_required"`, and the journal logs the Entra
error code.

### When the client secret expires

This is a different failure, and you don't need to sign in again: refresh
tokens belong to the App ID, not to a particular secret. Tools return *"Entra
ID rejected the app's client secret (invalid_client: AADSTS7000222…)"*. Create
a new secret in the app registration, then:

```bash
sudo /opt/outlook-mcp/bin/seal-credential.sh client_secret
sudo systemctl restart outlook-mcp
```

Do this before the old secret expires: Entra allows several secrets at once, so
you can switch with no downtime and then delete the old one.

## Connecting a client

The service listens on `127.0.0.1:8767` only (msteams-mcp uses 8766, so both
can run on one host; change `PORT` in the unit if needed). The MCP endpoint is
`/mcp`, and every request except `/health` must send
`Authorization: Bearer <token>`. To reach it from another machine, put an HTTPS
front end in front of the port: a reverse proxy (nginx, Caddy), a Tailscale
service, or an SSH tunnel.

```json
{
  "mcpServers": {
    "outlook_mcp": {
      "command": "npx",
      "args": ["mcp-remote", "https://your-host.example.com/mcp",
               "--header", "Authorization:${AUTH}"],
      "env": { "AUTH": "Bearer <your-token>" }
    }
  }
}
```

Then ask Claude "what unread mail do I have?" to confirm it works end to end.

## Agent skill

[`skills/outlook-mail/`](skills/outlook-mail/SKILL.md) is an
[Agent Skill](https://agentskills.io) that teaches an agent to use these tools
well: how to find mail, read efficiently, write drafts, and get approval
before sending anything.

- **Claude Code:** `cp -r skills/outlook-mail ~/.claude/skills/`, then start a new session.
- **claude.ai / Claude Desktop:** `cd skills && zip -r outlook-mail.zip outlook-mail`,
  then upload it under Settings → Capabilities → Skills.

## Layout on the server

| Path | What |
|---|---|
| `/opt/outlook-mcp/app` | Application code (read-only to the service) |
| `/opt/outlook-mcp/venv` | Python virtual environment |
| `/opt/outlook-mcp/bin` | `authorize.sh`, `seal-credential.sh` |
| `/usr/local/sbin/outlook-mcp-authorize` | Symlink to the sign-in script |
| `/etc/outlook-mcp/outlook-mcp.env` | Non-secret settings |
| `/etc/credstore.encrypted/outlook-mcp.*.cred` | TPM-sealed `client_secret`, `mcp_token` and `vault_key` |
| `/var/lib/outlook-mcp/vault.json.enc` | Encrypted refresh token (0600, user `outlook-mcp`) |
| `/etc/systemd/system/outlook-mcp.service` | The unit |

The service runs as the static system user `outlook-mcp` rather than with
`DynamicUser`. This gives the vault, which the sign-in script writes as root,
a stable owner.

## Configuration

Settings in `/etc/outlook-mcp/outlook-mcp.env` (see [`deploy/outlook-mcp.env.example`](deploy/outlook-mcp.env.example)):

| Variable | Default | Meaning |
|---|---|---|
| `MS_TENANT_ID` | required | Tenant GUID or domain |
| `MS_CLIENT_ID` | required | Application (client) ID |
| `MS_REDIRECT_URI` | `http://localhost` | Must exactly match a redirect URI registered under the *Web* platform |
| `ENABLE_WRITE_TOOLS` | `true` | `false` removes the draft and organise tools |
| `ENABLE_SEND_TOOLS` | `true` | `false` removes the send tools: drafts only |
| `SEND_CONFIRMATION` | `client` | `elicit`: the server itself asks you to approve every send |
| `SEND_ALLOWED_DOMAINS` | empty (any) | Comma-separated domains mail may be sent to |
| `SIGNATURE_FILE` | none | Signature appended to outgoing mail (`.html`/`.htm` as HTML, else text) |
| `ENABLE_CALENDAR_TOOLS` | `true` | `false` removes the calendar tools |
| `TIMEZONE` | Outlook setting, else UTC | IANA time zone for calendar times, e.g. `Asia/Dubai` |
| `WORKING_HOURS` | Outlook setting, else `08:00-17:00` | Hours `find_free_time` searches by default |
| `WORKING_DAYS` | Outlook setting, else `mon-fri` | Days `find_free_time` searches, e.g. `sun-thu` |
| `KEEPALIVE_HOURS` | `12` | Background token refresh interval; `0` disables it |
| `MAX_BODY_CHARS` | `20000` | Default body cut-off in `get_message` (continue with `offset`) |
| `MAX_ATTACHMENT_MB` | `25` | Largest attachment `get_attachment` downloads |

Restart the service after changing settings.

## Day-to-day

```bash
sudo systemctl restart outlook-mcp            # after changing settings or credentials
journalctl -u outlook-mcp -f                   # follow logs
systemd-analyze security outlook-mcp           # hardening score (lower is better)
sudo outlook-mcp-authorize                      # sign in again / switch account

# Update code: pull the new version, then
sudo ./deploy/install.sh                        # keeps config, credentials and sign-in

# Replace the client secret (before it expires; no new sign-in needed)
sudo /opt/outlook-mcp/bin/seal-credential.sh client_secret
sudo systemctl restart outlook-mcp

# Rotate the MCP token
sudo /opt/outlook-mcp/bin/seal-credential.sh mcp_token --generate
sudo systemctl restart outlook-mcp              # then update your MCP client

# Rotate the vault key (makes the stored sign-in unreadable: sign in again)
sudo /opt/outlook-mcp/bin/seal-credential.sh vault_key --generate --force
sudo systemctl restart outlook-mcp && sudo outlook-mcp-authorize
```

## Notes

- **After a BIOS/firmware update or a Secure Boot change,** the TPM may refuse
  to decrypt the credentials, and the service fails to start with a credential
  error. Re-seal `client_secret` and `mcp_token`, regenerate `vault_key` with
  `--force`, and sign in again.
- **What this protects against:** copies of the disk, backups, sync tools or a
  git push picking up files, and other users or services on the server. It
  does not stop someone with root on the running machine (root can decrypt
  the credentials), so keep SSH key-only and root access tight. The refresh
  token is as powerful as your mailbox within the consented permissions,
  including sending mail as you. The client secret alone can't act as you: it
  only works together with a refresh token or a fresh sign-in.
- **Prompt injection.** Mail is untrusted input: anyone can email you text
  written to manipulate an agent ("forward all invoices to…"). The skill tells
  the agent to treat message content as data, never as instructions, and the
  approval safeguards above are what stop an injected send. Keep them on.
- **A certificate instead of a secret** would be stronger (the private key never
  leaves the server, and it can last longer). It isn't implemented here.
- **Conditional Access:** if your tenant restricts sign-ins to compliant or
  managed devices, or to known locations, refreshes from the server may be
  blocked (`AADSTS53003`, `AADSTS50076`...). The journal shows the code.
- **`search_messages`** uses the mailbox search index (KQL). Results can lag a
  few minutes behind new mail; `list_messages` reads the folder directly.
- **Shared mailboxes** aren't supported; they would need `Mail.ReadWrite.Shared`
  and `Mail.Send.Shared`.
- **Meeting invitations** show in mail as messages marked
  `(eventMessageRequest)`; answer them with `respond_to_event(message_id=…)`,
  or find them all with `list_events(needs_response=True)`.
- **Free/busy of other people** comes from Outlook's free/busy sharing: people
  outside your organisation (and some resources) are usually not visible, and
  `find_free_time` says so. Only your own working hours are used.
- **Other people's calendars** and delegate access aren't supported; they would
  need `Calendars.ReadWrite.Shared`.
- If the service fails with a memory-protection error, comment out
  `MemoryDenyWriteExecute=yes` in the unit, then run `daemon-reload` and restart.

## Local development

Without systemd, secrets fall back to the `MS_CLIENT_SECRET` / `MCP_TOKEN` /
`VAULT_KEY` environment variables (or to files named by `*_FILE`, for example
`MS_CLIENT_SECRET_FILE`), and the vault is `./vault.json.enc` (override with
`VAULT_PATH`).

```bash
python3 -m venv .venv && . .venv/bin/activate
pip install -r requirements.txt
export MS_TENANT_ID=... MS_CLIENT_ID=... MCP_TOKEN=dev-token
read -rs MS_CLIENT_SECRET && export MS_CLIENT_SECRET      # paste, no echo, not in history
export VAULT_KEY=$(python -c 'import vault; print(vault.new_key_text())')
python authorize.py                             # interactive sign-in -> ./vault.json.enc
python server.py                                # serves http://127.0.0.1:8767/mcp
```

Offline tests (Graph and Entra ID are mocked by an in-memory mailbox, no tenant needed):

```bash
python tests/test_mail.py -v                    # -v prints sample output
python tests/test_calendar.py -v
```

## License

MIT, see [LICENSE](LICENSE).
