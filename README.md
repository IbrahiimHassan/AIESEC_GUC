# AIESEC GTa Opportunity Tracker

A Python automation tool that scrapes daily Global Talent (GTa) opportunities from AIESEC.org, detects newly opened listings, and emails them as formatted opportunity cards. It runs automatically every day on GitHub Actions.

## Why I built it
Our matching team at AIESEC GUC had to check the opportunities portal manually to spot new openings. This tool automates that, so new opportunities reach the team quickly and consistently, which speeds up applicant matching.

## How it works
1. **Scrape:** Selenium and BeautifulSoup collect the current opportunities: title, host organization, location, duration, number of applicants, and link.
2. **Store:** Results are saved to `Today.xlsx`. The previous run is kept as `Yesterday.xlsx`.
3. **Compare:** Today's and yesterday's data are compared to find new opportunities, which are written to `New.xlsx`.
4. **Notify:** New listings are sent by email through the Gmail API as HTML cards.
5. **Schedule:** A GitHub Actions workflow runs the script daily and commits the updated Excel files back to the repository.

## Example email
<img width="1499" height="695" alt="image" src="https://github.com/user-attachments/assets/ef920841-ca03-42fb-a35e-5b5087497f19" />


Each card shows the number of applicants, the opportunity title, the host organization, the location, the duration, and a link to view the opportunity.

## Tech stack
- Python
- Selenium and BeautifulSoup
- Excel (`.xlsx`) for daily snapshots
- Gmail API
- GitHub Actions

## Repository structure
```
AIESEC_GUC/
├── .github/workflows/   # scheduled daily run
├── aiesec.py            # scraping, comparison, and email logic
├── Today.xlsx           # latest scrape
├── Yesterday.xlsx       # previous scrape
├── New.xlsx             # newly detected opportunities
└── README.md
```

## Run it locally
1. Clone the repository:
```bash
   git clone https://github.com/IbrahiimHassan/AIESEC_GUC.git
   cd AIESEC_GUC
```
2. Install the dependencies:
```bash
   pip install selenium beautifulsoup4 pandas openpyxl google-api-python-client google-auth-oauthlib
```
3. Provide your Gmail API credentials and recipient list as environment variables.
4. Run:
```bash
   python aiesec.py
```

## Configuration
Gmail API credentials and recipient addresses are stored as GitHub Secrets and are not included in the code.

## Notes
Built for internal use at AIESEC GUC and shared as a portfolio project.
