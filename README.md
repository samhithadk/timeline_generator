# Carlsquare Timeline Generator — Web App

A Streamlit web app that generates M&A process timeline slides (.pptx) from a simple form.

## Files in this folder

| File | Purpose |
|------|---------|
| `app.py` | The Streamlit web app |
| `timeline_slide_generator.py` | The slide generation engine |
| `logo_light.png` | Coloured Carlsquare logo (light theme) |
| `logo_dark.png` | White Carlsquare logo (dark theme) |
| `requirements.txt` | Python dependencies for Streamlit Cloud |
| `.streamlit/config.toml` | App theme and server settings |

---

## How to deploy (one time, ~15 minutes)

### Step 1 — Put the files on GitHub

1. Go to [github.com](https://github.com) and sign in (or create a free account).
2. Click **New repository** → name it `carlsquare-timeline` → set to **Private** → click **Create repository**.
3. Upload all files in this folder to the repository (drag and drop works).

### Step 2 — Deploy on Streamlit Community Cloud

1. Go to [share.streamlit.io](https://share.streamlit.io) and sign in with your GitHub account.
2. Click **New app**.
3. Select your `carlsquare-timeline` repository.
4. Set **Main file path** to `app.py`.
5. Click **Deploy**.

That's it. Streamlit will install the dependencies and give you a live URL in about 2 minutes.

### Step 3 — Share the URL

Share the URL with your team. They'll be prompted for the password before they can use the app.

---

## Changing the password

Open `app.py` and change this line near the top:

```python
PASSWORD = "carlsquare2024"   # ← change this
```

Then save and push to GitHub — the app will update automatically within seconds.

---

## Updating the app

Any time you push a change to GitHub, Streamlit Cloud picks it up and redeploys automatically. No manual steps needed.
