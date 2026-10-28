# Tampermonkey Outlook scripts

Userscripts that make Outlook on the web nicer to use, especially when it's installed as a PWA (an app window in Edge or Chrome).

## Scripts

| Script | What it does | Install |
| --- | --- | --- |
| [Outlook PWA: unread badge + reload after sleep](scripts/outlook-pwa-badge.user.js) | Shows the Inbox unread count on the PWA's taskbar icon, and reloads Outlook after the laptop wakes from sleep so missed mail shows up. | [Install](https://raw.githubusercontent.com/Atsukio/tampermonkey-outlook-scripts/main/scripts/outlook-pwa-badge.user.js) |

## Installing

1. Install the [Tampermonkey](https://www.tampermonkey.net/) extension in Edge or Chrome.
2. Click the **Install** link for the script above. Tampermonkey opens an install page; click **Install**.
3. Reload Outlook (or reopen the Outlook PWA window).

Tampermonkey checks this repository for new versions automatically.

## Outlook PWA unread badge

Works on `outlook.office.com`, `outlook.office365.com`, `outlook.cloud.microsoft` and `outlook.live.com`.

- **Unread badge:** reads the Inbox unread count from the folder pane and shows it on the taskbar icon via the [Badging API](https://developer.mozilla.org/docs/Web/API/Badging_API). It updates every 10 seconds, and straight away when Outlook changes a folder's count. Outlook's own badge calls are blocked so it doesn't clear the badge when the window gets focus.
- **Reload after sleep:** if a timer tick is more than 5 minutes late, the machine has probably been asleep, so the page reloads to reconnect.

### Settings

Edit these constants at the top of the script (Tampermonkey dashboard → the script → Editor):

| Setting | Default | Meaning |
| --- | --- | --- |
| `FOLDER` | `'Inbox'` | Folder name to count. Change it if your Outlook isn't in English. |
| `POLL_MS` | 10 seconds | How often the badge is refreshed. |
| `SLEEP_GAP_MS` | 5 minutes | How late a timer tick has to be to count as a wake from sleep. |
| `DEBUG` | `true` | Logs to the browser console. Set to `false` once the badge works. |

### Troubleshooting

- **No badge:** the badge only shows in the installed PWA window, not a normal browser tab. Open DevTools in the PWA window (F12) and look for `[Outlook badge]` messages.
- **Badge doesn't change on Calendar or People:** the count is read from the mail folder pane, so the badge keeps its last value until you go back to Mail.

## License

[MIT](LICENSE)
