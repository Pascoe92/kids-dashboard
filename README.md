# 🌟 Kids Behaviour Dashboard

A gamified behaviour, chores & rewards dashboard for Myles (6) and Finley (3), built on Google Apps Script + Google Sheets.

## Features
- ✅ Daily chore checklist (kids tap their own chores)
- ⭐ Points system with progress toward rewards
- 🔥 Daily streak tracking with bonus points every 3 days
- ⚡ Strike system — 5 strikes = no treat (auto-resets every Monday)
- 🎁 Rewards shop with point redemption
- 🎉 Fireworks & confetti animations on achievements
- 🔐 Parent PIN zone for strikes, bonuses & reward management
- 📱 Mobile-friendly — works on any device

## Setup

### 1. Google Apps Script
1. Open [Google Sheets](https://docs.google.com/spreadsheets/d/1g1gohK6yKtyGPwdSd85tM-UvxhtaJsUhDGUOVq8RTfc)
2. Go to **Extensions → Apps Script**
3. Delete any existing code in `Code.gs` and paste in the contents of `Code.gs` from this repo
4. Create a new HTML file called `Index` (click the **+** next to Files → HTML) and paste in `Index.html`
5. Run `initializeSheets()` once to create all required tabs
6. **Deploy → New deployment** as a Web App:
   - Execute as: **Me**
   - Who has access: **Anyone**
7. Copy the deployment URL — that's your dashboard!

### 2. Change your parent PIN
- Open the **Config** tab in the Google Sheet
- Find the `parent_pin` row and change `1234` to your preferred 4-digit PIN

### 3. Customise rewards & chores
- Edit the **Rewards** tab to add/remove/change rewards and point costs
- Edit the **ChoresList** tab to add/remove chores (set Kids column to `Both`, `Myles`, or `Finley`)

### 4. Bookmark on devices
- Open the web app URL on each device and **Add to Home Screen** for app-like access

## Config options (Config sheet)
| Setting | Default | Description |
|---|---|---|
| `points_per_chore` | 10 | Points awarded per completed chore |
| `strike_limit` | 5 | Strikes before consequence triggers |
| `strike_consequence` | No ice cream! | Text shown when limit hit |
| `streak_bonus_days` | 3 | Days in a row to earn streak bonus |
| `streak_bonus_points` | 15 | Bonus points for a streak milestone |
| `parent_pin` | 1234 | PIN for the parent zone |

## Sheets
| Sheet | Purpose |
|---|---|
| `KidsData` | Live points, strikes, streak per kid |
| `ChoresList` | Available chores and which kids they apply to |
| `Rewards` | Reward catalog with point costs |
| `ChoresLog` | Full history of every chore and point transaction |
| `StrikesLog` | Full history of every strike |
| `Config` | App settings |
