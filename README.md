# NSE Stock Buy Signal Analysis

Automated daily NSE (National Stock Exchange) stock analysis using machine learning to predict buy signals. The system runs automatically via GitHub Actions and sends email notifications with top buy recommendations.

## Features

- **Machine Learning Analysis**: Uses XGBoost and Gradient Boosting classifiers to predict buy signals
- **Technical Indicators**: Analyzes multiple indicators including RSI, MACD, Bollinger Bands, Donchian Channels, ATR, and more
- **Automated Scheduling**: Runs automatically Monday-Friday at 15:40 IST (10:10 UTC) via GitHub Actions
- **Email Notifications**: Sends formatted HTML emails with top 30 buy signals and probabilities
- **Manual Execution**: Can be triggered manually from GitHub Actions tab

## Setup Instructions

### 1. Repository Structure

```
/your-repo
  ├─ PythonApplication1.py    # Main analysis script
  ├─ requirements.txt          # Python dependencies
  ├─ README.md                 # This file
  └─ .github/
      └─ workflows/
          └─ schedule.yml      # GitHub Actions workflow
```

### 2. Configure GitHub Secrets

Go to **Repository → Settings → Secrets and variables → Actions → New repository secret** and add:

#### Required Secrets:
- **SMTP_USER**: Your sender email address (e.g., `yourname@gmail.com` or `you@company.com`)
- **SMTP_PASS**: SMTP/App password for authentication
- **MAIL_RECIPIENTS**: Comma-separated list of recipient emails (e.g., `me@domain.com,team@domain.com`)

#### Email Provider Configuration:

**For Gmail:**
1. Go to Google Account → Security → App passwords
2. Create an App Password for "Mail"
3. Use this password for `SMTP_PASS`

**For Microsoft 365/Outlook:**
1. Ensure SMTP AUTH is enabled for the mailbox
2. Use your email password or app password if available
3. The script defaults to Office 365 SMTP settings (`smtp.office365.com:587`)

### 3. Schedule Configuration

The workflow runs **Monday-Friday at 15:40 IST (10:10 UTC)** by default.

To change the schedule, edit `.github/workflows/schedule.yml`:

```yaml
schedule:
  # Cron format: "minute hour day month weekday"
  - cron: "10 10 * * 1-5"  # 10:10 UTC Mon-Fri = 15:40 IST
```

**Note**: GitHub Actions uses UTC time. IST is UTC+5:30.

### 4. Manual Execution

You can manually trigger the workflow:
1. Go to **Actions** tab in your repository
2. Select **"Scheduled Buy Signals"** workflow
3. Click **"Run workflow"**
4. Select the branch and click **"Run workflow"**

## How It Works

### Analysis Process

1. **Data Collection**: Fetches 1 year of historical data for NSE stocks using yfinance
2. **Feature Engineering**: Calculates technical indicators and features
3. **Model Training**: Trains XGBoost classifier with SMOTE for handling imbalanced data
4. **Prediction**: Predicts buy signals for the next trading day
5. **Filtering**: Applies confidence thresholds and validation checks
6. **Ranking**: Ranks stocks by buy probability and selects top 30
7. **Notification**: Sends formatted email with results

### Key Features Used

- Donchian Channels (High, Low, Mid)
- Bollinger Bands (Upper, Lower)
- Simple Moving Average (SMA 20)
- ATR (Average True Range)
- RSI (Relative Strength Index)
- MACD and Signal Line
- On-Balance Volume (OBV)
- Momentum and Volatility
- Volume trends

### Thresholds

- **Buy Threshold**: 70% probability
- **Strong Buy Threshold**: 90% probability

## Email Format

The email includes:
- **Subject**: NSE Top Buy Signals – [Date]
- **Timestamp**: Generated time in IST
- **Table**: Rank, Ticker, Buy Probability, Today's Close, Predicted Class
- **Format**: Both plain text and HTML versions

## Requirements

All dependencies are listed in `requirements.txt`:
- pandas
- numpy
- yfinance
- scikit-learn
- xgboost
- imbalanced-learn
- scipy

Python 3.9+ required (for `zoneinfo` module, which is built-in).

## Customization

### Change Stock List

Edit the `nse_tickers` list in `PythonApplication1.py` to analyze different stocks.

### Modify Buy Thresholds

Adjust thresholds in `PythonApplication1.py`:

```python
strong_buy_threshold = 0.90  # For strong buy signals
buy_threshold = 0.70         # For regular buy signals
```

### Change SMTP Server

For non-Office 365 email providers, modify the function call in the script:

```python
send_top_buys_email(
    # ... other parameters ...
    smtp_server="smtp.gmail.com",  # Change SMTP server
    smtp_port=587,                 # Change port if needed
)
```

### Adjust Number of Results

Change the number of top signals (default is 30):

```python
top_buys = sorted(results, key=lambda x: x[1], reverse=True)[:30]  # Change 30 to desired number
```

## Weekends and Market Holidays

- The workflow is configured to run **Monday-Friday only** (1-5 in cron)
- The script will still execute on market holidays
- To skip market holidays, you can:
  - Use a holiday calendar package (e.g., `holidays`, `pandas-market-calendars`)
  - Check if today's data is current before sending emails

## Troubleshooting

### Workflow Not Running
- **Scheduled workflows only run on the default branch (main)**
- Check that secrets are properly configured
- Verify the cron schedule syntax
- Check GitHub Actions quota and permissions
- GitHub scheduled workflows may take up to 1 hour to first trigger after push
- You can manually test workflows using "Run workflow" button in Actions tab

### Email Not Sending
- Verify `SMTP_USER`, `SMTP_PASS`, and `MAIL_RECIPIENTS` are set
- Check email provider's SMTP settings
- Review workflow logs in Actions tab
- Ensure 2FA/App Passwords are configured for Gmail

### No Buy Signals
- This is normal when market conditions don't favor buy signals
- The script will print "No buy candidates were generated for this run"
- Email will still be sent with this message

### Python Version Issues
- Ensure Python 3.9+ is used (for `zoneinfo`)
- If using Python 3.8 or earlier, replace `zoneinfo` with `pytz`

## License

This project is for educational and personal use. Please ensure compliance with data provider terms of service and applicable regulations when using for trading decisions.

## Disclaimer

This tool is for informational purposes only. It does not constitute financial advice. Always do your own research and consult with financial advisors before making investment decisions. Past performance does not guarantee future results.
