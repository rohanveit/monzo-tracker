#!/usr/bin/env python3
"""Monzo API Transaction Scraper - Entry point."""

import argparse
import os

from dotenv import load_dotenv

from .auth import TokenManager, TOKEN_FILE
from .api import MonzoAPI
from .models import format_transaction
from .spreadsheet import write_transactions

# Load environment variables from .env file
load_dotenv()

# Configuration
CLIENT_ID = os.getenv('MONZO_CLIENT_ID')
CLIENT_SECRET = os.getenv('MONZO_CLIENT_SECRET')
REDIRECT_URI = os.getenv('MONZO_REDIRECT_URI', 'http://localhost:8080/callback')


def main():
    """Main execution flow."""
    parser = argparse.ArgumentParser(description="Monzo Transaction Tracker")
    parser.add_argument("--reauth", action="store_true", help="Force re-authentication with Monzo")
    args = parser.parse_args()

    print("=== Monzo API Transaction Scraper ===\n")

    if args.reauth:
        if TOKEN_FILE.exists():
            TOKEN_FILE.unlink()
            print("Tokens cleared. Will re-authenticate.\n")
        else:
            print("No saved tokens found.\n")

    # Check if credentials are set
    if not CLIENT_ID or not CLIENT_SECRET:
        print("ERROR: Please set your Monzo API credentials in the .env file!")
        print("\nYou need to:")
        print("1. Create a .env file in the same directory as this script")
        print("2. Add your credentials to the .env file:")
        print("   MONZO_CLIENT_ID=your_client_id_here")
        print("   MONZO_CLIENT_SECRET=your_client_secret_here")
        print("\nTo get credentials:")
        print("1. Go to https://developers.monzo.com/")
        print("2. Create an OAuth client")
        print("3. Set the redirect URI to: http://localhost:8080/callback")
        return

    try:
        # Initialize token manager and API client
        token_manager = TokenManager(CLIENT_ID, CLIENT_SECRET, REDIRECT_URI)
        api = MonzoAPI(token_manager)

        # Get accounts (will auto-authenticate if needed)
        accounts = api.get_accounts()
        print(f"Found {len(accounts)} account(s)")

        # Get transactions for each account
        for account in accounts:
            account_id = account['id']
            account_desc = account.get('description', 'Unknown')

            print(f"\n--- Account: {account_desc} ({account_id[:10]}...) ---")

            transactions = api.get_transactions(account_id, days=30)
            print(f"Retrieved {len(transactions)} transactions from last 30 days\n")

            if transactions:
                # Format and display transactions
                formatted_txs = [format_transaction(tx) for tx in transactions]

                # Sort by date (most recent first)
                formatted_txs.sort(key=lambda x: x.date, reverse=True)

                # Display transactions
                print(f"{'Date':<19} | {'Amount':>12} | {'Category':<15} | Description")
                print("-" * 80)
                for tx in formatted_txs:
                    print(f"{tx.date} | {tx.amount:>12} | {tx.category:<15} | {tx.description}")
                    if tx.notes:
                        print(f"               Notes: {tx.notes}")
                    print()

                # Save to spreadsheet
                filepath = write_transactions(formatted_txs)
                print(f"Transactions saved to {filepath}")

                # Summary
                total = sum(tx.amount_raw for tx in formatted_txs)
                print(f"\nTotal spent: £{total:.2f}")
            else:
                print("No transactions found in the last 30 days")

    except Exception as e:
        print(f"\nError: {str(e)}")
        import traceback
        traceback.print_exc()


if __name__ == '__main__':
    main()
