# Monzo Financial Tracker

A Python tool that connects to the Monzo API to fetch and track financial transactions.

## Setup

1. Install dependencies: `uv sync`
2. Create a `.env` file with your Monzo API credentials:
   ```
   MONZO_CLIENT_ID=your_client_id
   MONZO_CLIENT_SECRET=your_client_secret
   MONZO_REDIRECT_URI=http://localhost:8080/callback
   ```
3. Get credentials from https://developers.monzo.com/

## Usage

```bash
uv run monzo-tracker
```

On first run, a browser will open for Monzo authentication. Approve the login in your Monzo app within 5 minutes.

Tokens are saved to `~/.monzo_tokens.json` and automatically refreshed when expired.

If you would like to manually trigger re-auth, you can use:

```bash
uv run monzo-tracker --reauth
```

---

## TODO

- [ ] Integrate Ollama server for AI-powered transaction categorization
- [ ] Implement better/more descriptive category assignment using LLM
- [ ] Add spending analytics and summaries
- [ ] Create visualizations for spending patterns
- [ ] Dynamically generated formulae rather than relying on code to generate sums/averages
